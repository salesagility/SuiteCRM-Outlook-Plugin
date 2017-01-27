using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using SuiteCRMClient;
using SuiteCRMClient.Logging;
using SuiteCRMClient.RESTObjects;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    public class AppointmentSyncing: Syncing
    {
        List<AppointmentSyncState> lCalItems;

        public AppointmentSyncing(SyncContext context)
            : base(context)
        {
        }

        public void StartCalendarSync()
        {
            try
            {
                Log.Info("AppointmentSync thread started");
                Outlook.NameSpace oNS = this.Application.GetNamespace("mapi");
                if (oNS.Categories["SuiteCRM"] == null)
                {
                    oNS.Categories.Add("SuiteCRM", Outlook.OlCategoryColor.olCategoryColorGreen, Outlook.OlCategoryShortcutKey.olCategoryShortcutKeyNone);
                }
                Outlook.MAPIFolder appointmentsFolder = GetDefaultFolder("appointments");
                Outlook.Items items = appointmentsFolder.Items;

                items.ItemAdd -= Items_ItemAdd;
                items.ItemChange -= Items_ItemChange;
                items.ItemRemove -= Items_ItemRemove;
                items.ItemAdd += Items_ItemAdd;
                items.ItemChange += Items_ItemChange;
                items.ItemRemove += Items_ItemRemove;

                GetOutlookCalItems(appointmentsFolder);
                SyncMeetings(appointmentsFolder, "Meetings");
                SyncMeetings(appointmentsFolder, "Calls");
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.StartCalendarSync", ex);
            }
            finally
            {
                Log.Info("AppointmentSync thread completed");
            }
        }

        void Items_ItemRemove()
        {
            try
            {
                if (IsCalendarView)
                {
                    try
                    {
                        foreach (var oItem in lCalItems)
                        {
                            try
                            {
                                string sID = oItem.OutlookItem.EntryID;
                            }
                            catch (COMException)
                            {
                                eNameValue[] data = new eNameValue[2];
                                data[0] = clsSuiteCRMHelper.SetNameValuePair("id", oItem.SEntryID);
                                data[1] = clsSuiteCRMHelper.SetNameValuePair("deleted", "1");
                                clsSuiteCRMHelper.SetEntryUnsafe(data, oItem.SType);
                                oItem.Delete = true;
                            }
                        }
                        lCalItems.RemoveAll(a => a.Delete);
                    }
                    catch
                    { }
                }
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.Items_ItemRemove", ex);
            }
        }

        void Items_ItemChange(object Item)
        {
            Log.Warn("Items_ItemChange");
            try
            {
                var aItem = Item as Outlook.AppointmentItem;

                string entryId = aItem.EntryID;
                AppointmentSyncState callitem = lCalItems.FirstOrDefault(a => a.OutlookItem.EntryID == entryId);
                Log.Warn("CalItem EntryID=  " + aItem.EntryID);
                if (callitem != default(AppointmentSyncState))
                {
                    var utcNow = DateTime.UtcNow;
                    if (Math.Abs((int)(utcNow - callitem.OModifiedDate).TotalSeconds) > 5)
                    {
                        Log.Warn("2 callitem.IsUpdate = " + callitem.IsUpdate);
                        callitem.IsUpdate = 0;
                    }

                    Log.Warn("Before UtcNow - callitem.OModifiedDate= " + (int)(utcNow - callitem.OModifiedDate).TotalSeconds);

                    if (Math.Abs((int)(utcNow - callitem.OModifiedDate).TotalSeconds) > 2 && callitem.IsUpdate == 0)
                    {
                        callitem.OModifiedDate = DateTime.UtcNow;
                        Log.Warn("1 callitem.IsUpdate = " + callitem.IsUpdate);
                        callitem.IsUpdate++;
                    }

                    Log.Warn("callitem = " + callitem.OutlookItem.Subject);
                    Log.Warn("callitem.SEntryID = " + callitem.SEntryID);
                    Log.Warn("callitem mod_date= " + callitem.OModifiedDate.ToString());
                    Log.Warn("utcNow= " + DateTime.UtcNow.ToString());
                    Log.Warn("UtcNow - callitem.OModifiedDate= " + (int)(DateTime.UtcNow - callitem.OModifiedDate).TotalSeconds);
                }
                else
                {
                    Log.Warn("not found callitem ");
                }


                if (IsCalendarView && lCalItems.Exists(a => a.OutlookItem.EntryID == aItem.EntryID
                                 && callitem.IsUpdate == 1
                                 )
                )
                {
                    Outlook.UserProperty oProp = aItem.UserProperties["SType"];
                    Outlook.UserProperty oProp1 = aItem.UserProperties["SEntryID"];
                    if (oProp != null && oProp1 != null)
                    {
                        callitem.IsUpdate++;
                        AddAppointmentToS(aItem, oProp.Value.ToString(), oProp1.Value.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.Items_ItemChange", ex);
            }
        }

        void Items_ItemAdd(object Item)
        {
            Log.Warn("Items_ItemAdd");
            var aItem = Item as Outlook.AppointmentItem;
            if (IsCalendarView && !lCalItems.Exists(a => a.OutlookItem.EntryID == aItem.EntryID))
            {
                AddAppointmentToS(aItem, "Meetings");
            }
        }

        private void GetOutlookCalItems(Outlook.MAPIFolder appointmentsFolder)
        {
            try
            {
                if (lCalItems == null)
                {
                    lCalItems = new List<AppointmentSyncState>();
                    Outlook.Items items = appointmentsFolder.Items; //.Restrict("[MessageClass] = 'IPM.Appointment'" + GetStartDateString());
                    foreach (Outlook.AppointmentItem aItem in items)
                    {
                        if (aItem.Start < DateTime.Now.AddDays(-5))
                            continue;
                        Outlook.UserProperty oProp = aItem.UserProperties["SOModifiedDate"];
                        if (oProp != null)
                        {
                            Outlook.UserProperty oProp1 = aItem.UserProperties["SType"];
                            Outlook.UserProperty oProp2 = aItem.UserProperties["SEntryID"];
                            lCalItems.Add(new AppointmentSyncState
                            {
                                OutlookItem = aItem,
                                OModifiedDate = DateTime.UtcNow,
                                SType = oProp1.Value.ToString(),
                                SEntryID = oProp2.Value.ToString()
                            });
                        }
                        else
                        {
                            lCalItems.Add(new AppointmentSyncState
                            {
                                OutlookItem = aItem,
                                SType = "Meetings"
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.GetOutlookCalItems", ex);
            }
        }

        private void SetRecepients(Outlook.AppointmentItem aItem, string sMeetingID, string sModule)
        {
            aItem.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
            int iCount = aItem.Recipients.Count;
            for (int iItr = 1; iItr <= iCount; iItr++)
            {
                aItem.Recipients.Remove(1);
            }

            eEntryValue[] Users;
            string[] invitee_categories = { "users", "contacts", "leads" };
            foreach (string invitee_category in invitee_categories)
            {
                Users = clsSuiteCRMHelper.getRelationships(sModule, sMeetingID, invitee_category, new string[] { "id", "email1", "phone_work" });
                if (Users != null)
                {

                    foreach (var oResult1 in Users)
                    {
                        dynamic dResult1 = JsonConvert.DeserializeObject(oResult1.name_value_object.ToString());

                        Log.Warn("-------------------SetRecepients-----Start-----dResult1---2-------");
                        Log.Warn((string)Convert.ToString(dResult1));
                        Log.Warn("-------------------SetRecepients-----End---------------");

                        string phone_work = dResult1.phone_work.value.ToString();
                        string sTemp =
                            (sModule == "Meetings") || String.IsNullOrEmpty(phone_work) || String.IsNullOrWhiteSpace(phone_work) ?
                                dResult1.email1.value.ToString() :
                                dResult1.email1.value.ToString() + ":" + phone_work;
                        aItem.Recipients.Add(sTemp);
                    }
                }
            }
        }

        private void SetMeetings(eEntryValue[] el, Outlook.MAPIFolder appointmentsFolder, string sModule)
        {

            foreach (var oResult in el)
            {
                try
                {
                    UpdateFromCrm(appointmentsFolder, sModule, oResult);
                }
                catch (Exception ex)
                {
                    Log.Error("ThisAddIn.SyncMeetings", ex);
                }
            }
        }

        private void UpdateFromCrm(Outlook.MAPIFolder appointmentsFolder, string sModule, eEntryValue oResult)
        {
            dynamic dResult = JsonConvert.DeserializeObject(oResult.name_value_object.ToString());
            DateTime date_start = DateTime.ParseExact(dResult.date_start.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);
            date_start = date_start.Add(new DateTimeOffset(DateTime.Now).Offset);
            if (date_start < GetStartDate())
            {
                return;
            }

            var oItem = lCalItems.FirstOrDefault(a => a.SEntryID == dResult.id.value.ToString() && a.SType == sModule);
            if (oItem == default(AppointmentSyncState))
            {
                Outlook.AppointmentItem aItem = appointmentsFolder.Items.Add(Outlook.OlItemType.olAppointmentItem);
                aItem.Subject = dResult.name.value.ToString();
                aItem.Body = dResult.description.value.ToString();
                if (!string.IsNullOrWhiteSpace(dResult.date_start.value.ToString()))
                {
                    aItem.Start = date_start;
                    int iMin = 0, iHour = 0;
                    if (!string.IsNullOrWhiteSpace(dResult.duration_minutes.value.ToString()))
                    {
                        iMin = int.Parse(dResult.duration_minutes.value.ToString());
                    }
                    if (!string.IsNullOrWhiteSpace(dResult.duration_hours.value.ToString()))
                    {
                        iHour = int.Parse(dResult.duration_hours.value.ToString());
                    }
                    if (sModule == "Meetings")
                    {
                        aItem.Location = dResult.location.value.ToString();
                        aItem.End = aItem.Start;
                        if (iHour > 0)
                            aItem.End.AddHours(iHour);
                        if (iMin > 0)
                            aItem.End.AddMinutes(iMin);
                    }
                    Log.Warn("   default SetRecepients");
                    SetRecepients(aItem, dResult.id.value.ToString(), sModule);

                    //}
                    try
                    {
                        aItem.Duration = iMin + iHour*60;
                    }
                    catch (Exception)
                    {
                    }
                }
                Outlook.UserProperty oProp = aItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                oProp.Value = dResult.date_modified.value.ToString();
                Outlook.UserProperty oProp1 = aItem.UserProperties.Add("SType", Outlook.OlUserPropertyType.olText);
                oProp1.Value = sModule;
                Outlook.UserProperty oProp2 = aItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                oProp2.Value = dResult.id.value.ToString();
                lCalItems.Add(new AppointmentSyncState
                {
                    OutlookItem = aItem,
                    OModifiedDate = DateTime.ParseExact(dResult.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null),
                    SType = sModule,
                    SEntryID = dResult.id.value.ToString(),
                    Touched = true
                });
                aItem.Save();
            }
            else
            {
                oItem.Touched = true;
                Outlook.AppointmentItem aItem = oItem.OutlookItem;
                Outlook.UserProperty oProp = aItem.UserProperties["SOModifiedDate"];

                if (oProp.Value != dResult.date_modified.value.ToString())
                {
                    aItem.Subject = dResult.name.value.ToString();
                    aItem.Body = dResult.description.value.ToString();
                    if (!string.IsNullOrWhiteSpace(dResult.date_start.value.ToString()))
                    {
                        aItem.Start = date_start;
                        int iMin = 0, iHour = 0;
                        if (!string.IsNullOrWhiteSpace(dResult.duration_minutes.value.ToString()))
                        {
                            iMin = int.Parse(dResult.duration_minutes.value.ToString());
                        }
                        if (!string.IsNullOrWhiteSpace(dResult.duration_hours.value.ToString()))
                        {
                            iHour = int.Parse(dResult.duration_hours.value.ToString());
                        }
                        if (sModule == "Meetings")
                        {
                            aItem.Location = dResult.location.value.ToString();
                            aItem.End = aItem.Start;
                            if (iHour > 0)
                                aItem.End.AddHours(iHour);
                            if (iMin > 0)
                                aItem.End.AddMinutes(iMin);
                            Log.Warn("    SetRecepients");
                            SetRecepients(aItem, dResult.id.value.ToString(), sModule);
                        }
                        try
                        {
                            aItem.Duration = iMin + iHour*60;
                        }
                        catch (Exception)
                        {
                        }
                    }

                    if (oProp == null)
                        oProp = aItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                    oProp.Value = dResult.date_modified.value.ToString();
                    Outlook.UserProperty oProp1 = aItem.UserProperties["SType"];
                    if (oProp1 == null)
                        oProp1 = aItem.UserProperties.Add("SType", Outlook.OlUserPropertyType.olText);
                    oProp1.Value = sModule;
                    Outlook.UserProperty oProp2 = aItem.UserProperties["SEntryID"];
                    if (oProp2 == null)
                        oProp2 = aItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                    oProp2.Value = dResult.id.value.ToString();
                    aItem.Save();
                }
                Log.Warn((string) ("Not default dResult.date_modified= " + dResult.date_modified.value.ToString()));
                oItem.OModifiedDate = DateTime.ParseExact(dResult.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);
            }
        }

        private void SyncMeetings(Outlook.MAPIFolder appointmentsFolder, string sModule)
        {
            Log.Warn("SyncMeetings");
            try
            {
                int iOffset = 0;
                while (true)
                {
                    eGetEntryListResult _result2 = clsSuiteCRMHelper.GetEntryList(sModule, "assigned_user_id = '" + clsSuiteCRMHelper.GetUserId() + "'",
                                    0, "date_start DESC", iOffset, false, clsSuiteCRMHelper.GetSugarFields(sModule));

                    var nextOffset = _result2.next_offset;
                    if (iOffset == nextOffset)
                        break;

                    SetMeetings(_result2.entry_list, appointmentsFolder, sModule);

                    iOffset = nextOffset;
                    if (iOffset == 0)
                        break;
                }
                eEntryValue[] invited = clsSuiteCRMHelper.getRelationships("Users", clsSuiteCRMHelper.GetUserId(), sModule.ToLower(), clsSuiteCRMHelper.GetSugarFields(sModule));
                if (invited != null)
                {
                    SetMeetings(invited, appointmentsFolder, sModule);
                }

                try
                {
                    if (sModule == "Meetings")
                    {
                        var lItemToBeDeletedO = lCalItems.Where(a => !a.Touched && !string.IsNullOrWhiteSpace(a.OModifiedDate.ToString()) && a.SType == sModule);
                        foreach (var oItem in lItemToBeDeletedO)
                        {
                            try
                            {
                                oItem.OutlookItem.Delete();
                            }
                            catch (Exception)
                            {
                                Log.Warn("   Exception  oItem.oItem.Delete");
                            }
                        }
                        lCalItems.RemoveAll(a => !a.Touched && !string.IsNullOrWhiteSpace(a.OModifiedDate.ToString()) && a.SType == sModule);
                    }
                    var lItemToBeAddedToS = lCalItems.Where(a => !a.Touched && string.IsNullOrWhiteSpace(a.OModifiedDate.ToString()) && a.SType == sModule);
                    foreach (var oItem in lItemToBeAddedToS)
                    {
                        AddAppointmentToS(oItem.OutlookItem, sModule);
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("ThisAddIn.SyncMeetings", ex);
                }
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.SyncMeetings", ex);
            }
        }

        private void AddAppointmentToS(Outlook.AppointmentItem aItem, string sModule, string sID = "")
        {
            Log.Warn("AddAppointmentToS");
            if (!settings.SyncCalendar)
                return;
            if (aItem != null)
            {
                try
                {
                    string _result = "";
                    eNameValue[] data = new eNameValue[8];
                    DateTime uTCDateTime = new DateTime();
                    DateTime time2 = new DateTime();
                    uTCDateTime = aItem.Start.ToUniversalTime();
                    time2 = aItem.End.ToUniversalTime();
                    string str = string.Format("{0:yyyy-MM-dd HH:mm:ss}", uTCDateTime);
                    string str2 = string.Format("{0:yyyy-MM-dd HH:mm:ss}", time2);
                    int num = aItem.Duration / 60;
                    int num2 = aItem.Duration % 60;
                    data[0] = clsSuiteCRMHelper.SetNameValuePair("name", aItem.Subject);
                    data[1] = clsSuiteCRMHelper.SetNameValuePair("description", aItem.Body);
                    data[2] = clsSuiteCRMHelper.SetNameValuePair("location", aItem.Location);
                    data[3] = clsSuiteCRMHelper.SetNameValuePair("date_start", str);
                    data[4] = clsSuiteCRMHelper.SetNameValuePair("date_end", str2);
                    data[5] = clsSuiteCRMHelper.SetNameValuePair("duration_minutes", num2.ToString());
                    data[6] = clsSuiteCRMHelper.SetNameValuePair("duration_hours", num.ToString());
                    if (sID == "")
                        data[7] = clsSuiteCRMHelper.SetNameValuePair("assigned_user_id", clsSuiteCRMHelper.GetUserId());
                    else
                        data[7] = clsSuiteCRMHelper.SetNameValuePair("id", sID);

                    _result = clsSuiteCRMHelper.SetEntryUnsafe(data, sModule);
                    if (sID == "")
                    {
                        Log.Warn("    -- AddAppointmentToS AddAppointmentToS sID =" + sID);

                        eSetRelationshipValue info = new eSetRelationshipValue
                        {
                            module2 = "meetings",
                            module2_id = _result,
                            module1 = "Users",
                            module1_id = clsSuiteCRMHelper.GetUserId()
                        };
                        clsSuiteCRMHelper.SetRelationshipUnsafe(info);

                    }
                    if (aItem.Recipients != null)
                    {
                        foreach (Outlook.Recipient objRecepient in aItem.Recipients)
                        {
                            try
                            {
                                Log.Warn("objRecepientName= " + objRecepient.Name.ToString());
                                Log.Warn("objRecepient= " + objRecepient.Address.ToString());
                            }
                            catch
                            {
                                Log.Warn("objRecepient ERROR");
                                continue;
                            }

                            string sCID = GetID(objRecepient.Address, "Contacts");
                            if (sCID != "")
                            {
                                eSetRelationshipValue info = new eSetRelationshipValue
                                {
                                    module2 = "meetings",
                                    module2_id = _result,
                                    module1 = "Contacts",
                                    module1_id = sCID
                                };

                                Log.Warn("    SetRelationship 1");
                                Log.Warn("    sCID=" + sCID);
                                clsSuiteCRMHelper.SetRelationshipUnsafe(info);

                                string AccountID = clsSuiteCRMHelper.getRelationship("Contacts", sCID, "accounts");

                                if (AccountID != "")
                                {
                                    info = new eSetRelationshipValue
                                    {
                                        module2 = "meetings",
                                        module2_id = _result,
                                        module1 = "Accounts",
                                        module1_id = AccountID
                                    };
                                    clsSuiteCRMHelper.SetRelationshipUnsafe(info);
                                }
                                continue;
                            }
                            sCID = GetID(objRecepient.Address, "Users");
                            if (sCID != "")
                            {
                                eSetRelationshipValue info = new eSetRelationshipValue
                                {
                                    module2 = "meetings",
                                    module2_id = _result,
                                    module1 = "Users",
                                    module1_id = sCID
                                };
                                clsSuiteCRMHelper.SetRelationshipUnsafe(info);
                                continue;
                            }
                            sCID = GetID(objRecepient.Address, "Leads");
                            if (sCID != "")
                            {
                                eSetRelationshipValue info = new eSetRelationshipValue
                                {
                                    module2 = "meetings",
                                    module2_id = _result,
                                    module1 = "Leads",
                                    module1_id = sCID
                                };
                                Log.Warn("    SetRelationship 2");
                                clsSuiteCRMHelper.SetRelationshipUnsafe(info);
                                continue;
                            }
                        }
                    }
                    Outlook.UserProperty oProp = aItem.UserProperties["SOModifiedDate"];
                    if (oProp == null)
                        oProp = aItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                    oProp.Value = DateTime.UtcNow;
                    Outlook.UserProperty oProp1 = aItem.UserProperties["SType"];
                    if (oProp1 == null)
                        oProp1 = aItem.UserProperties.Add("SType", Outlook.OlUserPropertyType.olText);
                    oProp1.Value = sModule;
                    Outlook.UserProperty oProp2 = aItem.UserProperties["SEntryID"];
                    if (oProp2 == null)
                        oProp2 = aItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                    oProp2.Value = _result;
                    Log.Warn("    AddAppointmentToS Save ");
                    aItem.Save();
                    string entryId = aItem.EntryID;
                    var sItem = lCalItems.FirstOrDefault(a => a.OutlookItem.EntryID == entryId);
                    if (sItem != default(AppointmentSyncState))
                    {
                        sItem.OutlookItem = aItem;
                        sItem.OModifiedDate = DateTime.UtcNow;
                        sItem.SEntryID = _result;
                        Log.Warn("    AddAppointmentToS Edit ");
                    }
                    else
                    {
                        lCalItems.Add(new AppointmentSyncState { SEntryID = _result, SType = sModule, OModifiedDate = DateTime.UtcNow, OutlookItem = aItem });
                        Log.Warn("    AddAppointmentToS New ");
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("ThisAddIn.AddAppointmentToS", ex);
                }
            }
        }

        public string GetID(string sEmailID, string sModule)
        {
            string str5 = "(" + sModule.ToLower() + ".id in (select eabr.bean_id from email_addr_bean_rel eabr INNER JOIN email_addresses ea on eabr.email_address_id = ea.id where eabr.bean_module = '" + sModule + "' and ea.email_address LIKE '%" + SuiteCRMAddIn.clsGlobals.MySqlEscape(sEmailID) + "%'))";

            Log.Warn("-------------------GetID-----Start---------------");

            Log.Warn("    str5=" + str5);

            Log.Warn("-------------------GetID-----End---------------");

            string[] fields = new string[1];
            fields[0] = "id";
            eGetEntryListResult _result = clsSuiteCRMHelper.GetEntryList(sModule, str5, settings.SyncMaxRecords, "date_entered DESC", 0, false, fields);
            if (_result.result_count > 0)
            {
                return clsSuiteCRMHelper.GetValueByKey(_result.entry_list[0], "id");
            }
            return "";
        }

        public Outlook.MAPIFolder GetDefaultFolder(string type)
        {
            return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
        }

        protected bool IsCalendarView => Context.CurrentFolderItemType == Outlook.OlItemType.olAppointmentItem;
    }
}
