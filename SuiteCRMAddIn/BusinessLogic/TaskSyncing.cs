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
        /// <summary>
        /// The module I synchronise with.
        /// </summary>
        const string CrmModule = "Tasks";

        public TaskSyncing(string name, SyncContext context)
            : base(name, context)
        {
            this.fetchQueryPrefix = string.Empty;
        }

        protected override string DefaultCrmModule
        {
            get
            {
                return CrmModule;
            }
        }

        public override bool SyncingEnabled => settings.SyncCalendar;

        public override void SynchroniseAll()
        {
            Outlook.NameSpace oNS = this.Application.GetNamespace("mapi");
            Outlook.MAPIFolder folder = GetDefaultFolder();

            GetOutlookItems(folder);
            SyncFolder(folder, CrmModule);
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

        protected override void SaveItem(Outlook.TaskItem olItem)
        {
            olItem.Save();
        }

        /// <summary>
        /// Synchronise items in the specified folder with the specified SuiteCRM module.
        /// </summary>
        /// <param name="folder">The folder.</param>
        /// <param name="crmModule">The module to snychronise with.</param>
        protected override void SyncFolder(Outlook.MAPIFolder folder, string crmModule)
        {
            Log.Info($"ContactSyncing.SyncFolder: '{folder}'");
            try
            {
                var untouched = new HashSet<SyncState<Outlook.TaskItem>>(ItemsSyncState);

                MergeRecordsFromCrm(folder, crmModule, untouched);

                try
                {
                    ResolveUnmatchedItems(untouched);
                }
                catch (Exception ex)
                {
                    Log.Error("TaskSyncing.SyncFolder", ex);
                }
            }
            catch (Exception ex)
            {
                Log.Error("TaskSyncing.SyncFolder", ex);
            }
        }


        /// <summary>
        /// Log a message regarding this Outlook appointment.
        /// </summary>
        /// <param name="olItem">The outlook item.</param>
        /// <param name="message">The message to be logged.</param>
        protected override void LogItemAction(Outlook.TaskItem olItem, string message)
        {
            try
            {
                Outlook.UserProperty olPropertyEntryId = olItem.UserProperties["SEntryID"];
                string crmId = olPropertyEntryId == null ?
                    "[not present]" :
                    olPropertyEntryId.Value;
                Log.Info($"{0}:\n\tOutlook Id  : {olItem.EntryID}\n\tCRM Id      : {crmId}\n\tSubject    : '{olItem.Subject}'\n\tStatus      : {olItem.Status}");
            }
            catch (COMException)
            {
                // Ignore: happens if the outlook item is already deleted.
            }
        }


        protected override SyncState<Outlook.TaskItem> UpdateFromCrm(Outlook.MAPIFolder tasksFolder, string crmType, eEntryValue crmItem)
        {
            SyncState<Outlook.TaskItem> result = null;

            if (clsSuiteCRMHelper.GetUserId() == crmItem.GetValueAsString("assigned_user_id"))
            {

                DateTime? date_start = null;
                DateTime? date_due = null;

                string time_start = "--:--", time_due = "--:--";

                if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("date_start")))
                {
                    Log.Warn("\tSET date_start = dResult.date_start");
                    date_start = DateTime.ParseExact(crmItem.GetValueAsString("date_start"), "yyyy-MM-dd HH:mm:ss", null);

                    date_start = date_start.Value.Add(new DateTimeOffset(DateTime.Now).Offset);
                    time_start =
                        TimeSpan.FromHours(date_start.Value.Hour)
                            .Add(TimeSpan.FromMinutes(date_start.Value.Minute))
                            .ToString(@"hh\:mm");
                }

                if (date_start != null && date_start >= GetStartDate())
                {

                    if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("date_due")))
                    {
                        date_due = DateTime.ParseExact(crmItem.GetValueAsString("date_due"), "yyyy-MM-dd HH:mm:ss", null);
                        date_due = date_due.Value.Add(new DateTimeOffset(DateTime.Now).Offset);
                        time_due =
                            TimeSpan.FromHours(date_due.Value.Hour).Add(TimeSpan.FromMinutes(date_due.Value.Minute)).ToString(@"hh\:mm");
                        ;
                    }

                    var oItem = ItemsSyncState.FirstOrDefault(a => a.CrmEntryId == crmItem.GetValueAsString("id"));

                    if (oItem == null)
                    {
                        result = AddNewItemFromCrmToOutlook(tasksFolder, crmItem, date_start, date_due, time_start, time_due);
                    }
                    else
                    {
                        result = UpdateExistingOutlookItemFromCrm(crmItem, date_start, date_due, time_start, time_due, oItem);
                    }
                }
            }

            return result;
        }

        private SyncState<Outlook.TaskItem> UpdateExistingOutlookItemFromCrm(eEntryValue crmItem, DateTime? date_start, DateTime? date_due, string time_start, string time_due, SyncState<Outlook.TaskItem> syncStateForItem)
        {
            Outlook.TaskItem outlookItem = syncStateForItem.OutlookItem;
            Outlook.UserProperty oProp = outlookItem.UserProperties["SOModifiedDate"];

            if (oProp.Value != crmItem.GetValueAsString("date_modified"))
            {
                SetOutlookItemPropertiesFromCrmItem(crmItem, date_start, date_due, time_start, time_due, outlookItem);
                outlookItem.Save();
            }
            syncStateForItem.OModifiedDate = DateTime.ParseExact(crmItem.GetValueAsString("date_modified"), "yyyy-MM-dd HH:mm:ss", null);

            return syncStateForItem;
        }

        private void SetOutlookItemPropertiesFromCrmItem(eEntryValue crmItem, DateTime? date_start, DateTime? date_due, string time_start, string time_due, Outlook.TaskItem outlookItem)
        {
            outlookItem.Subject = crmItem.GetValueAsString("name");

            if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("date_start")))
            {
                Log.Warn("\ttItem.StartDate= " + outlookItem.StartDate + ", date_start=" + date_start);
                outlookItem.StartDate = date_start.Value;
            }
            if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("date_due")))
            {
                outlookItem.DueDate = date_due.Value; // DateTime.Parse(dResult.date_due.value.ToString());
            }

            string body = crmItem.GetValueAsString("description");
            outlookItem.Body = string.Concat(body, "#<", time_start, "#", time_due);
            outlookItem.Status = GetStatus(crmItem.GetValueAsString("status"));
            outlookItem.Importance = GetImportance(crmItem.GetValueAsString("priority"));
            EnsureSynchronisationPropertiesForOutlookItem(outlookItem, crmItem.GetValueAsString("date_modified"), DefaultCrmModule, crmItem.id);
        }

        private SyncState<Outlook.TaskItem> AddNewItemFromCrmToOutlook(Outlook.MAPIFolder tasksFolder, eEntryValue crmItem, DateTime? date_start, DateTime? date_due, string time_start, string time_due)
        {
            Outlook.TaskItem olItem = tasksFolder.Items.Add(Outlook.OlItemType.olTaskItem);
            this.SetOutlookItemPropertiesFromCrmItem(crmItem, date_start, date_due, time_start, time_due, olItem);

            var newState = new TaskSyncState
            {
                OutlookItem = olItem,
                OModifiedDate = DateTime.ParseExact(crmItem.GetValueAsString("date_modified"), "yyyy-MM-dd HH:mm:ss", null),
                CrmEntryId = crmItem.GetValueAsString("id"),
            };
            ItemsSyncState.Add(newState);
            olItem.Save();
            LogItemAction(olItem, "AppointmentSyncing.AddNewItemFromCrmToOutlook");

            return newState;
        }

        protected override void GetOutlookItems(Outlook.MAPIFolder taskFolder)
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
                        AddOrUpdateItemFromOutlookToCrm(oItem, oProp1.Value.ToString());
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
                        AddOrUpdateItemFromOutlookToCrm(item, this.DefaultCrmModule, oProp2.Value);
                    }
                    else
                    {
                        AddOrUpdateItemFromOutlookToCrm(item, this.DefaultCrmModule);
                    }
                }
        }

        protected override void EnsureSynchronisationPropertyForOutlookItem(Outlook.TaskItem olItem, string name, string value)
        {
            Outlook.UserProperty olProperty = olItem.UserProperties[name];
            if (olProperty == null)
            {
                olProperty = olItem.UserProperties.Add(name, Outlook.OlUserPropertyType.olText);
            }
            olProperty.Value = value;
        }

        /// <summary>
        /// Construct a JSON packet representing this Outlook item, and despatch it to CRM. 
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmType">The type within CRM to which the item should be added.</param>
        /// <param name="entryId">The corresponding entry id in CRM, if known.</param>
        /// <returns>The CRM id of the object created or modified.</returns>
        protected override string ConstructAndDespatchCrmItem(Outlook.TaskItem olItem, string crmType, string entryId)
        {
            return clsSuiteCRMHelper.SetEntryUnsafe(new ProtoTask(olItem).AsNameValues(entryId), crmType);
        }

        public override Outlook.MAPIFolder GetDefaultFolder()
        {
            return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks);
        }

        protected override SyncState<Outlook.TaskItem> ConstructSyncState(Outlook.TaskItem oItem)
        {
            return new TaskSyncState
            {
                OutlookItem = oItem,
                CrmEntryId = oItem.UserProperties["SEntryID"]?.Value.ToString(),
                OModifiedDate = ParseDateTimeFromUserProperty(oItem.UserProperties["SOModifiedDate"]?.Value.ToString()),
            };
        }

        protected override SyncState<Outlook.TaskItem> GetExistingSyncState(Outlook.TaskItem oItem)
        {
            return ItemsSyncState.FirstOrDefault(a => a.OutlookItem.EntryID == oItem.EntryID);
        }

        protected override bool IsCurrentView => Context.CurrentFolderItemType == Outlook.OlItemType.olTaskItem;

        // Should presumably be removed at some point. Existing code was ignoring deletions for Contacts and Tasks
        // (but not for Appointments).
        protected override bool PropagatesLocalDeletions => false;

        /// <summary>
        /// Broadly, a C# representation of a CRM task.
        /// </summary>
        private class ProtoTask
        {
            private Outlook.TaskItem oItem;
            private string dateStart = string.Empty, dateDue = string.Empty;
            private string body = String.Empty;
            private string description = String.Empty;

            public string Body
            {
                get
                {
                    return body;
                }
            }

            public string DateStart
            {
                get
                {
                    return dateStart;
                }
            }

            public string DateDue
            {
                get
                {
                    return dateDue;
                }
            }

            public string Description
            {
                get
                {
                    return description;
                }
            }

            public string Priority
            {
                get
                {
                    string result;
                    switch (oItem.Importance)
                    {
                        case Outlook.OlImportance.olImportanceLow:
                            result = "Low";
                            break;

                        case Outlook.OlImportance.olImportanceNormal:
                            result = "Medium";
                            break;

                        case Outlook.OlImportance.olImportanceHigh:
                            result = "High";
                            break;
                        default:
                            result = string.Empty;
                            break;
                    }

                    return result;
                }
            }

            public string Status
            {
                get
                {
                    string result;
                    switch (oItem.Status)
                    {
                        case Outlook.OlTaskStatus.olTaskNotStarted:
                            result = "Not Started";
                            break;
                        case Outlook.OlTaskStatus.olTaskInProgress:
                            result = "In Progress";
                            break;
                        case Outlook.OlTaskStatus.olTaskComplete:
                            result = "Completed";
                            break;
                        case Outlook.OlTaskStatus.olTaskDeferred:
                            result = "Deferred";
                            break;
                        default:
                            result = string.Empty;
                            break;
                    }

                    return result;
                }
            }

            public ProtoTask(Outlook.TaskItem oItem)
            {
                this.oItem = oItem;
                DateTime uTCDateTime = new DateTime();
                DateTime time2 = new DateTime();
                uTCDateTime = oItem.StartDate.ToUniversalTime();
                if (oItem.DueDate != null)
                    time2 = oItem.DueDate.ToUniversalTime();

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
                            dateStart = string.Format("{0:yyyy-MM-dd HH:mm:ss}", uTCDateTime.ToUniversalTime());
                        if (time2.ToUniversalTime().Year < 4000)
                            dateDue = string.Format("{0:yyyy-MM-dd HH:mm:ss}", time2.ToUniversalTime());
                    }
                    else
                    {
                        dateStart = oItem.StartDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
                        dateDue = oItem.DueDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
                    }

                }
                else
                {
                    dateStart = oItem.StartDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
                    dateDue = oItem.DueDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
                }

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
            }

            /// <summary>
            /// Construct a name value list (to be serialised as JSON) representing this task.
            /// </summary>
            /// <param name="entryId">The presumed id of this task in CRM, if known.</param>
            /// <returns>a name value list representing this task</returns>
            public List<eNameValue> AsNameValues(string entryId)
            {
                var dataList = new List<eNameValue>();
                dataList.Add(clsSuiteCRMHelper.SetNameValuePair("name", this.oItem.Subject));
                dataList.Add(clsSuiteCRMHelper.SetNameValuePair("description", this.Description));
                dataList.Add(clsSuiteCRMHelper.SetNameValuePair("status", this.Status));
                dataList.Add(clsSuiteCRMHelper.SetNameValuePair("date_due", this.DateDue));
                dataList.Add(clsSuiteCRMHelper.SetNameValuePair("date_start", this.DateStart));
                dataList.Add(clsSuiteCRMHelper.SetNameValuePair("priority", this.Priority));

                dataList.Add(String.IsNullOrEmpty(entryId) ?
                    clsSuiteCRMHelper.SetNameValuePair("assigned_user_id", clsSuiteCRMHelper.GetUserId()) :
                    clsSuiteCRMHelper.SetNameValuePair("id", entryId));
                return dataList;
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
                    // Log.Warn("Body doesn't have time string");
                    return null;
                }
            }
        }
    }
}
