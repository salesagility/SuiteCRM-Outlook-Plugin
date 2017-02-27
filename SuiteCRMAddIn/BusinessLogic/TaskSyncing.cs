using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using SuiteCRMClient.Logging;
using SuiteCRMClient;
using SuiteCRMClient.RESTObjects;
using Newtonsoft.Json;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    public class TaskSyncing: Synchroniser<Outlook.TaskItem>
    {
        public TaskSyncing(SyncContext context)
            : base("Task synchroniser", context)
        {
        }

        public override bool SyncingEnabled => settings.SyncCalendar;

        public override void SynchroniseAll()
        {
            Outlook.NameSpace oNS = this.Application.GetNamespace("mapi");
            Outlook.MAPIFolder folder = GetDefaultFolder();

            GetOutlookItems(folder);
            SyncFolder(folder);
        }

        private Outlook.OlImportance GetImportance(string sImportance)
        {
            Outlook.OlImportance oPriority = Outlook.OlImportance.olImportanceLow;
            switch (sImportance)
            {
                case "High":
                    oPriority = Outlook.OlImportance.olImportanceHigh;
                    break;
                case "Medium":
                    oPriority = Outlook.OlImportance.olImportanceNormal;
                    break;
            }
            return oPriority;
        }
        private Outlook.OlTaskStatus GetStatus(string sStatus)
        {
            Outlook.OlTaskStatus oStatus = Outlook.OlTaskStatus.olTaskNotStarted;
            switch (sStatus)
            {
                case "In Progress":
                    oStatus = Outlook.OlTaskStatus.olTaskInProgress;
                    break;
                case "Completed":
                    oStatus = Outlook.OlTaskStatus.olTaskComplete;
                    break;
                case "Deferred":
                    oStatus = Outlook.OlTaskStatus.olTaskDeferred;
                    break;

            }
            return oStatus;
        }
        private void SyncFolder(Outlook.MAPIFolder tasksFolder)
        {
            Log.Warn("SyncTasks");
            Log.Warn("My UserId= " + clsSuiteCRMHelper.GetUserId());
            try
            {
                var untouched = new HashSet<SyncState<Outlook.TaskItem>>(ItemsSyncState);
                int iOffset = 0;
                while (true)
                {
                    eGetEntryListResult _result2 = clsSuiteCRMHelper.GetEntryList("Tasks", String.Empty,
                                    0, "date_start DESC", iOffset, false, clsSuiteCRMHelper.GetSugarFields("Tasks"));
                    var nextOffset = _result2.next_offset;
                    if (iOffset == nextOffset)
                        break;

                    foreach (var oResult in _result2.entry_list)
                    {
                        try
                        {
                            var state = UpdateFromCrm(tasksFolder, oResult);
                            if (state != null) untouched.Remove(state);
                        }
                        catch (Exception ex)
                        {
                            Log.Error("ThisAddIn.SyncTasks", ex);
                        }
                    }

                    iOffset = nextOffset;
                    if (iOffset == 0)
                        break;
                }
                try
                {
                    var lItemToBeDeletedO = untouched.Where(a => a.ExistedInCrm);
                    foreach (var oItem in lItemToBeDeletedO)
                    {
                        oItem.OutlookItem.Delete();
                        ItemsSyncState.Remove(oItem);
                    }

                    var lItemToBeAddedToS = untouched.Where(a => !a.ExistedInCrm);
                    foreach (var oItem in lItemToBeAddedToS)
                    {
                        AddToCrm(oItem.OutlookItem);
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("ThisAddIn.SyncTasks", ex);
                }
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.SyncTasks", ex);
            }
        }

        private SyncState<Outlook.TaskItem> UpdateFromCrm(Outlook.MAPIFolder tasksFolder, eEntryValue oResult)
        {
            dynamic dResult = JsonConvert.DeserializeObject(oResult.name_value_object.ToString());
            //
            if (clsSuiteCRMHelper.GetUserId() != dResult.assigned_user_id.value.ToString())
                return null;

            DateTime? date_start = null;
            DateTime? date_due = null;

            string time_start = "--:--", time_due = "--:--";


            if (!string.IsNullOrWhiteSpace(dResult.date_start.value.ToString()) &&
                !string.IsNullOrEmpty(dResult.date_start.value.ToString()))
            {
                Log.Warn("\tSET date_start = dResult.date_start");
                date_start = DateTime.ParseExact(dResult.date_start.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);

                date_start = date_start.Value.Add(new DateTimeOffset(DateTime.Now).Offset);
                time_start =
                    TimeSpan.FromHours(date_start.Value.Hour)
                        .Add(TimeSpan.FromMinutes(date_start.Value.Minute))
                        .ToString(@"hh\:mm");
            }

            if (date_start != null && date_start < GetStartDate())
            {
                Log.Warn("\tdate_start=" + date_start.ToString() + ", GetStartDate= " + GetStartDate().ToString());
                return null;
            }

            if (!string.IsNullOrWhiteSpace(dResult.date_due.value.ToString()))
            {
                date_due = DateTime.ParseExact(dResult.date_due.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);
                date_due = date_due.Value.Add(new DateTimeOffset(DateTime.Now).Offset);
                time_due =
                    TimeSpan.FromHours(date_due.Value.Hour).Add(TimeSpan.FromMinutes(date_due.Value.Minute)).ToString(@"hh\:mm");
                ;
            }

            foreach (var lt in ItemsSyncState)
            {
                Log.Warn("\tTask= " + lt.CrmEntryId);
            }

            var oItem = ItemsSyncState.FirstOrDefault(a => a.CrmEntryId == dResult.id.value.ToString());

            if (oItem == null)
            {
                Log.Warn("\tif default");
                Outlook.TaskItem tItem = tasksFolder.Items.Add(Outlook.OlItemType.olTaskItem);
                tItem.Subject = dResult.name.value.ToString();

                if (!string.IsNullOrWhiteSpace(dResult.date_start.value.ToString()))
                {
                    tItem.StartDate = date_start.Value;
                }
                if (!string.IsNullOrWhiteSpace(dResult.date_due.value.ToString()))
                {
                    tItem.DueDate = date_due.Value; // DateTime.Parse(dResult.date_due.value.ToString());
                }

                string body = dResult.description.value.ToString();
                tItem.Body = string.Concat(body, "#<", time_start, "#", time_due);
                tItem.Status = GetStatus(dResult.status.value.ToString());
                tItem.Importance = GetImportance(dResult.priority.value.ToString());

                Outlook.UserProperty oProp = tItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                oProp.Value = dResult.date_modified.value.ToString();
                Outlook.UserProperty oProp2 = tItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                oProp2.Value = dResult.id.value.ToString();
                var newState = new TaskSyncState
                {
                    OutlookItem = tItem,
                    OModifiedDate = DateTime.ParseExact(dResult.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null),
                    CrmEntryId = dResult.id.value.ToString(),
                };
                ItemsSyncState.Add(newState);
                Log.Warn("\tsave 0");
                tItem.Save();
                return newState;
            }
            else
            {
                Log.Warn("\telse not default");
                Outlook.TaskItem tItem = oItem.OutlookItem;
                Outlook.UserProperty oProp = tItem.UserProperties["SOModifiedDate"];

                Log.Warn(
                    (string)
                    ("\toProp.Value= " + oProp.Value + ", dResult.date_modified=" + dResult.date_modified.value.ToString()));
                if (oProp.Value != dResult.date_modified.value.ToString())
                {
                    tItem.Subject = dResult.name.value.ToString();

                    if (!string.IsNullOrWhiteSpace(dResult.date_start.value.ToString()))
                    {
                        Log.Warn("\ttItem.StartDate= " + tItem.StartDate + ", date_start=" + date_start);
                        tItem.StartDate = date_start.Value;
                    }
                    if (!string.IsNullOrWhiteSpace(dResult.date_due.value.ToString()))
                    {
                        tItem.DueDate = date_due.Value; // DateTime.Parse(dResult.date_due.value.ToString());
                    }

                    string body = dResult.description.value.ToString();
                    tItem.Body = string.Concat(body, "#<", time_start, "#", time_due);
                    tItem.Status = GetStatus(dResult.status.value.ToString());
                    tItem.Importance = GetImportance(dResult.priority.value.ToString());
                    if (oProp == null)
                        oProp = tItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                    oProp.Value = dResult.date_modified.value.ToString();
                    Outlook.UserProperty oProp2 = tItem.UserProperties["SEntryID"];
                    if (oProp2 == null)
                        oProp2 = tItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                    oProp2.Value = dResult.id.value.ToString();
                    Log.Warn("\tsave 1");
                    tItem.Save();
                }
                oItem.OModifiedDate = DateTime.ParseExact(dResult.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);
                return oItem;
            }
        }

        private void GetOutlookItems(Outlook.MAPIFolder taskFolder)
        {
            try
            {
                if (ItemsSyncState == null)
                {
                    ItemsSyncState = new ThreadSafeList<SyncState<Outlook.TaskItem>>();
                    Outlook.Items items = taskFolder.Items; //.Restrict("[MessageClass] = 'IPM.Task'" + GetStartDateString());
                    foreach (Outlook.TaskItem oItem in items)
                    {
                        if (oItem.DueDate < DateTime.Now.AddDays(-5))
                            continue;
                        Outlook.UserProperty oProp = oItem.UserProperties["SOModifiedDate"];
                        if (oProp != null)
                        {
                            Outlook.UserProperty oProp2 = oItem.UserProperties["SEntryID"];
                            ItemsSyncState.Add(new TaskSyncState
                            {
                                OutlookItem = oItem,
                                //OModifiedDate = "Fresh",
                                OModifiedDate = DateTime.UtcNow,

                                CrmEntryId = oProp2.Value.ToString()
                            });
                        }
                        else
                        {
                            ItemsSyncState.Add(new TaskSyncState
                            {
                                OutlookItem = oItem
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.GetOutlookTItems", ex);
            }
        }

        override protected void OutlookItemChanged(Outlook.TaskItem oItem)
        {
            Log.Debug("Outlook Tasks ItemChange");
                string entryId = oItem.EntryID;
                Log.Warn("\toItem.EntryID= " + entryId);

                var taskitem = ItemsSyncState.FirstOrDefault(a => a.OutlookItem.EntryID == entryId);
                if (taskitem != null)
                {
                    if ((DateTime.UtcNow - taskitem.OModifiedDate).TotalSeconds > 5)
                    {
                        Log.Warn("2 callitem.IsUpdate = " + taskitem.IsUpdate);
                        taskitem.IsUpdate = 0;
                    }

                    Log.Warn("Before UtcNow - callitem.OModifiedDate= " + (DateTime.UtcNow - taskitem.OModifiedDate).TotalSeconds.ToString());

                    if ((int)(DateTime.UtcNow - taskitem.OModifiedDate).TotalSeconds > 2 && taskitem.IsUpdate == 0)
                    {
                        taskitem.OModifiedDate = DateTime.UtcNow;
                        Log.Warn("1 callitem.IsUpdate = " + taskitem.IsUpdate);
                        taskitem.IsUpdate++;
                    }

                    Log.Warn("callitem = " + taskitem.OutlookItem.Subject);
                    Log.Warn("callitem.SEntryID = " + taskitem.CrmEntryId);
                    Log.Warn("callitem mod_date= " + taskitem.OModifiedDate.ToString());
                    Log.Warn("UtcNow - callitem.OModifiedDate= " + (DateTime.UtcNow - taskitem.OModifiedDate).TotalSeconds.ToString());
                }
                else
                {
                    Log.Warn("not found callitem ");
                }


                if (IsCurrentView && ItemsSyncState.Exists(a => a.OutlookItem.EntryID == entryId //// if (IsTaskView && lTaskItems.Exists(a => a.oItem.EntryID == entryId && a.OModifiedDate != "Fresh"))
                                 && taskitem.IsUpdate == 1
                                 )
                )
                {

                    Outlook.UserProperty oProp1 = oItem.UserProperties["SEntryID"];
                    if (oProp1 != null)
                    {
                        Log.Warn("\tgo to AddTaskToS");
                        taskitem.IsUpdate++;
                        AddToCrm(oItem, oProp1.Value.ToString());
                    }
                }
        }

        override protected void OutlookItemAdded(Outlook.TaskItem item)
        {
                if (IsCurrentView)
                {
                    Outlook.UserProperty oProp2 = item.UserProperties["SEntryID"];  // to avoid duplicating of the task
                    if (oProp2 != null)
                    {
                        AddToCrm(item, oProp2.Value);
                    }
                    else
                    {
                        AddToCrm(item);
                    }
                }
        }

        private void AddToCrm(Outlook.TaskItem oItem, string sID = "")
        {
            Log.Warn("AddTaskToS");
            //if (!settings.SyncCalendar)
            //    return;
            if (oItem == null) return;
            try
            {
                string _result = String.Empty;
                eNameValue[] data = new eNameValue[7];
                string strStatus = String.Empty;
                string strImportance = String.Empty;
                switch (oItem.Status)
                {
                    case Outlook.OlTaskStatus.olTaskNotStarted:
                        strStatus = "Not Started";
                        break;
                    case Outlook.OlTaskStatus.olTaskInProgress:
                        strStatus = "In Progress";
                        break;
                    case Outlook.OlTaskStatus.olTaskComplete:
                        strStatus = "Completed";
                        break;
                    case Outlook.OlTaskStatus.olTaskDeferred:
                        strStatus = "Deferred";
                        break;
                }
                switch (oItem.Importance)
                {
                    case Outlook.OlImportance.olImportanceLow:
                        strImportance = "Low";
                        break;

                    case Outlook.OlImportance.olImportanceNormal:
                        strImportance = "Medium";
                        break;

                    case Outlook.OlImportance.olImportanceHigh:
                        strImportance = "High";
                        break;
                }

                DateTime uTCDateTime = new DateTime();
                DateTime time2 = new DateTime();
                uTCDateTime = oItem.StartDate.ToUniversalTime();
                if (oItem.DueDate != null)
                    time2 = oItem.DueDate.ToUniversalTime();

                string body = String.Empty;
                string str, str2;
                str = str2 = String.Empty;
                if (oItem.Body != null)
                {
                    body = oItem.Body.ToString();
                    var times = this.ParseTimesFromTaskBody(body);
                    if (times != null)
                    {
                        uTCDateTime = uTCDateTime.Add(times[0]);
                        time2 = time2.Add(times[1]);

                        //check max date, date must has value !
                        if (uTCDateTime.ToUniversalTime().Year < 4000)
                            str = string.Format("{0:yyyy-MM-dd HH:mm:ss}", uTCDateTime.ToUniversalTime());
                        if (time2.ToUniversalTime().Year < 4000)
                            str2 = string.Format("{0:yyyy-MM-dd HH:mm:ss}", time2.ToUniversalTime());
                    }
                    else
                    {
                        str = oItem.StartDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
                        str2 = oItem.DueDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
                    }

                }
                else
                {
                    str = oItem.StartDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
                    str2 = oItem.DueDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
                }

                //str = "2016-11-10 11:34:01";
                //str2 = "2016-11-19 11:34:01";


                string description = String.Empty;

                if (!string.IsNullOrEmpty(body))
                {
                    int lastIndex = body.LastIndexOf("#<");
                    if (lastIndex >= 0)
                        description = body.Remove(lastIndex);
                    else
                    {
                        description = body;
                    }
                }
                Log.Warn("\tdescription= " + description);

                data[0] = clsSuiteCRMHelper.SetNameValuePair("name", oItem.Subject);
                data[1] = clsSuiteCRMHelper.SetNameValuePair("description", description);
                data[2] = clsSuiteCRMHelper.SetNameValuePair("status", strStatus);
                data[3] = clsSuiteCRMHelper.SetNameValuePair("date_due", str2);
                data[4] = clsSuiteCRMHelper.SetNameValuePair("date_start", str);
                data[5] = clsSuiteCRMHelper.SetNameValuePair("priority", strImportance);

                if (sID == String.Empty)
                    data[6] = clsSuiteCRMHelper.SetNameValuePair("assigned_user_id", clsSuiteCRMHelper.GetUserId());
                else
                    data[6] = clsSuiteCRMHelper.SetNameValuePair("id", sID);

                _result = clsSuiteCRMHelper.SetEntryUnsafe(data, "Tasks");
                Outlook.UserProperty oProp = oItem.UserProperties["SOModifiedDate"];
                if (oProp == null)
                    oProp = oItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                oProp.Value = DateTime.UtcNow;
                Outlook.UserProperty oProp2 = oItem.UserProperties["SEntryID"];
                if (oProp2 == null)
                    oProp2 = oItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                oProp2.Value = _result;
                string entryId = oItem.EntryID;
                oItem.Save();

                var sItem = ItemsSyncState.FirstOrDefault(a => a.OutlookItem.EntryID == entryId);
                if (sItem != null)
                {
                    sItem.OutlookItem = oItem;
                    sItem.OModifiedDate = DateTime.UtcNow;
                    sItem.CrmEntryId = _result;
                }
                else
                    ItemsSyncState.Add(new TaskSyncState { CrmEntryId = _result, OModifiedDate = DateTime.UtcNow, OutlookItem = oItem });

                Log.Warn("\tdate_start= " + str + ", date_due=" + str2);
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.AddTaskToS", ex);
            }
        }

        private TimeSpan[] ParseTimesFromTaskBody(string body)
        {
            try
            {
                if (string.IsNullOrEmpty(body))
                    return null;
                TimeSpan[] timesToAdd = new TimeSpan[2];
                List<int> hhmm = new List<int>(4);

                string times = body.Substring(body.LastIndexOf("#<")).Substring(2);
                char[] sep = { '<', '#', ':' };
                int parsed = 0;
                foreach (var digit in times.Split(sep))
                {
                    int.TryParse(digit, out parsed);
                    hhmm.Add(parsed);
                    parsed = 0;
                }

                TimeSpan start_time = TimeSpan.FromHours(hhmm[0]).Add(TimeSpan.FromMinutes(hhmm[1]));
                TimeSpan due_time = TimeSpan.FromHours(hhmm[2]).Add(TimeSpan.FromMinutes(hhmm[3]));
                timesToAdd[0] = start_time;
                timesToAdd[1] = due_time;
                return timesToAdd;
            }
            catch
            {
                Log.Warn("Body doesn't have time string");
                return null;
            }
        }

        public override Outlook.MAPIFolder GetDefaultFolder()
        {
            return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks);
        }

        protected override bool IsCurrentView => Context.CurrentFolderItemType == Outlook.OlItemType.olTaskItem;

        // Should presumably be removed at some point. Existing code was ignoring deletions for Contacts and Tasks
        // (but not for Appointments).
        protected override bool PropagatesLocalDeletions => false;
    }
}
