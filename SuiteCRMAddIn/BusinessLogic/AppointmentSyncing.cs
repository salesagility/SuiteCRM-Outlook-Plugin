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
    public class AppointmentSyncing: Syncing<Outlook.AppointmentItem>
    {
        public AppointmentSyncing(SyncContext context)
            : base(context)
        {
        }

        public override bool SyncingEnabled => settings.SyncCalendar;

        public void StartSync()
        {
            try
            {
                Log.Info("AppointmentSync thread started");
                AddSuiteCrmOutlookCategory();
                Outlook.MAPIFolder folder = GetDefaultFolder();

                GetOutlookItems(folder);
                SyncFolder(folder, "Meetings");
                SyncFolder(folder, "Calls");
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

        // TODO: Should _not_ be here. This category is used by all Syncing classes and email archiving,
        // so should be added near add-in start-up.
        private void AddSuiteCrmOutlookCategory()
        {
            Outlook.NameSpace oNS = this.Application.GetNamespace("mapi");
            if (oNS.Categories["SuiteCRM"] == null)
            {
                oNS.Categories.Add("SuiteCRM", Outlook.OlCategoryColor.olCategoryColorGreen,
                    Outlook.OlCategoryShortcutKey.olCategoryShortcutKeyNone);
            }
        }

        override protected void OutlookItemChanged(Outlook.AppointmentItem aItem)
        {
            try
            {
                string entryId = aItem.EntryID;
                var callitem = ItemsSyncState.FirstOrDefault(a => a.OutlookItem.EntryID == entryId);
                Log.Warn("CalItem EntryID=  " + aItem.EntryID);
                if (callitem != null)
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
                    Log.Warn("callitem.SEntryID = " + callitem.CrmEntryId);
                    Log.Warn("callitem mod_date= " + callitem.OModifiedDate.ToString());
                    Log.Warn("utcNow= " + DateTime.UtcNow.ToString());
                    Log.Warn("UtcNow - callitem.OModifiedDate= " + (int)(DateTime.UtcNow - callitem.OModifiedDate).TotalSeconds);
                }
                else
                {
                    Log.Warn("not found callitem ");
                }


                if (IsCurrentView && ItemsSyncState.Exists(a => a.OutlookItem.EntryID == aItem.EntryID
                                 && callitem.IsUpdate == 1
                                 )
                )
                {
                    Outlook.UserProperty olPropertyType = aItem.UserProperties["SType"];
                    Outlook.UserProperty olPropertyEntryId = aItem.UserProperties["SEntryID"];
                    if (olPropertyType != null && olPropertyEntryId != null)
                    {
                        callitem.IsUpdate++;
                        AddItemFromOutlookToCrm(aItem, olPropertyType.Value.ToString(), olPropertyEntryId.Value.ToString());
                    }
                }
            }
            finally
            {
            }
        }

        /// <summary>
        /// Entry point from event handler when an item is added in Outlook.
        /// </summary>
        /// <param name="appointment"></param>
        override protected void OutlookItemAdded(Outlook.AppointmentItem appointment)
        {
            if (IsCurrentView && !this.ItemsSyncState.Exists(a => a.OutlookItem.EntryID == appointment.EntryID))
            {
                AddItemFromOutlookToCrm(appointment, "Meetings");
            }
        }

        /// <summary>
        /// Get all items in this appointments folder
        /// </summary>
        /// <param name="appointmentsFolder"></param>
        private void GetOutlookItems(Outlook.MAPIFolder appointmentsFolder)
        {
            try
            {
                if (ItemsSyncState == null)
                {
                    ItemsSyncState = new List<SyncState<Outlook.AppointmentItem>>();

                    foreach (Outlook.AppointmentItem aItem in appointmentsFolder.Items)
                    {
                        if (aItem.Start >= this.GetStartDate())
                        {
                            Outlook.UserProperty olPropertyModified = aItem.UserProperties["SOModifiedDate"];
                            if (olPropertyModified != null)
                            {
                                /* The appointment probably already has the three magic properties 
                                 * required for synchronisation; is that a proxy for believing that it
                                 * already exists in CRM? */
                                Outlook.UserProperty olPropertyType = aItem.UserProperties["SType"];
                                Outlook.UserProperty olPropertyEntryId = aItem.UserProperties["SEntryID"];
                                var crmType = olPropertyType.Value.ToString();
                                ItemsSyncState.Add(new AppointmentSyncState(crmType)
                                {
                                    OutlookItem = aItem,
                                    OModifiedDate = DateTime.UtcNow,
                                    CrmEntryId = olPropertyEntryId.Value.ToString()
                                });
                            }
                            else
                            {
                                ItemsSyncState.Add(new AppointmentSyncState("Meetings")
                                {
                                    OutlookItem = aItem,
                                });
                            }
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

            string[] invitee_categories = { "users", "contacts", "leads" };
            foreach (string invitee_category in invitee_categories)
            {
                eEntryValue[] Users = clsSuiteCRMHelper.getRelationships(sModule, sMeetingID, invitee_category, new string[] { "id", "email1", "phone_work" });
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

        /// <summary>
        /// Update these appointments 
        /// </summary>
        /// <param name="appointments">The meetings to be synchronised.</param>
        /// <param name="appointmentsFolder">The outlook folder to synchronise into.</param>
        /// <param name="crmType">The type of CRM objects represented by the appointments.</param>
        /// <param name="untouched">A list of items which have not yet been synchronised; this list is 
        /// modified (destructuvely changed) by the action of this method.</param>
        private void UpdateAppointmentsFromCrmToOutlook(eEntryValue[] appointments, Outlook.MAPIFolder appointmentsFolder, string crmType, HashSet<SyncState<Outlook.AppointmentItem>> untouched)
        {

            foreach (var appointment in appointments)
            {
                try
                {
                    var state = MaybeUpdateAppointmentFromCrmToOutlook(appointmentsFolder, crmType, appointment);
                    if (state != null)
                    {
                        // i.e., the entry was updated...
                        untouched.Remove(state);
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("AppointmentSyncing.UpdateAppointmentsFromCrm", ex);
                }
            }
        }

        /// <summary>
        /// Update a single appointment in the specified Outlook folder with changes from CRM, but 
        /// only if its start date is fewer than five days in the past.
        /// </summary>
        /// <param name="appointmentsFolder">The folder to synchronise into.</param>
        /// <param name="crmType">The CRM type of the candidate item.</param>
        /// <param name="candidateItem">The candidate item from CRM.</param>
        /// <returns>The synchronisation state of the item updated (if it was updated).</returns>
        private SyncState<Outlook.AppointmentItem> MaybeUpdateAppointmentFromCrmToOutlook(
            Outlook.MAPIFolder appointmentsFolder, 
            string crmType, 
            eEntryValue candidateItem)
        {
            SyncState<Outlook.AppointmentItem> result = null;
            dynamic crmItem = JsonConvert.DeserializeObject(candidateItem.name_value_object.ToString());
            DateTime date_start = DateTime.ParseExact(crmItem.date_start.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);
            date_start = date_start.Add(new DateTimeOffset(DateTime.Now).Offset); // correct for offset from UTC.
            if (date_start >= GetStartDate())
            {
                /* search for the item among the items I already know about */
                var oItem = this.ItemsSyncState.FirstOrDefault(a => a.CrmEntryId == crmItem.id.value.ToString() && a.CrmType == crmType);
                if (oItem == null)
                {
                    /* didn't find it, so add it to Outlook */
                    result = AddNewItemFromCrmToOutlook(appointmentsFolder, crmType, crmItem, date_start);
                }
                else
                {
                    /* found it, so update it from the CRM item */
                    result = UpdateExistingOutlookItemFromCrm(crmType, crmItem, date_start, oItem);
                }
            }

            return result;
        }

        /// <summary>
        /// Update an existing Outlook item with values taken from a corresponding CRM item. Note that 
        /// this just overwrites all values in the Outlook item.
        /// </summary>
        /// <param name="crmType">The CRM type of the item from which values are to be taken.</param>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <param name="date_start">The state date/time of the item, adjusted for timezone.</param>
        /// <param name="oItem">The outlook item assumed to correspond with the CRM item.</param>
        /// <returns></returns>
        private SyncState<Outlook.AppointmentItem> UpdateExistingOutlookItemFromCrm(string crmType, dynamic crmItem, DateTime date_start, SyncState<Outlook.AppointmentItem> oItem)
        {
            Outlook.AppointmentItem olAppointment = oItem.OutlookItem;
            Outlook.UserProperty olPropertyModifiedDate = olAppointment.UserProperties["SOModifiedDate"];

            if (olPropertyModifiedDate.Value != crmItem.date_modified.value.ToString())
            {
                olAppointment.Subject = crmItem.name.value.ToString();
                olAppointment.Body = crmItem.description.value.ToString();
                if (!string.IsNullOrWhiteSpace(crmItem.date_start.value.ToString()))
                {
                    olAppointment.Start = date_start;
                    int iMin = 0, iHour = 0;
                    if (!string.IsNullOrWhiteSpace(crmItem.duration_minutes.value.ToString()))
                    {
                        iMin = int.Parse(crmItem.duration_minutes.value.ToString());
                    }
                    if (!string.IsNullOrWhiteSpace(crmItem.duration_hours.value.ToString()))
                    {
                        iHour = int.Parse(crmItem.duration_hours.value.ToString());
                    }
                    /* TODO: Why only meetings? Is this bug #7? */
                    if (crmType == "Meetings")
                    {
                        olAppointment.Location = crmItem.location.value.ToString();
                        olAppointment.End = olAppointment.Start;
                        if (iHour > 0)
                            olAppointment.End.AddHours(iHour);
                        if (iMin > 0)
                            olAppointment.End.AddMinutes(iMin);
                        Log.Warn("    SetRecepients");
                        SetRecepients(olAppointment, crmItem.id.value.ToString(), crmType);
                    }
                    olAppointment.Duration = iMin + iHour * 60;
                }

                SetupSynchronisationPropertiesForOutlookItem(olAppointment, crmType, crmItem);
                olAppointment.Save();
            }
            Log.Warn((string)("Not default dResult.date_modified= " + crmItem.date_modified.value.ToString()));
            oItem.OModifiedDate = DateTime.ParseExact(crmItem.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);

            return oItem;
        }

        /// <summary>
        /// There are a set of properties which are essential for synchronisation. Ensure this item has them.
        /// TODO: Possibly a candidate for refactoring to superclass.
        /// </summary>
        /// <param name="olItem">The Outlook item to be synchronised.</param>
        /// <param name="crmType">The CRM type of the object to synchronise with.</param>
        /// <param name="crmAppointment">The CRM object to synchronise with.</param>
        private static void SetupSynchronisationPropertiesForOutlookItem(Outlook.AppointmentItem olItem, string crmType, dynamic crmAppointment)
        {
            EnsureSynchronisationPropertiesForOutlookItem(olItem, crmAppointment.date_modified.value.ToString(), crmType, crmAppointment.id.value.ToString());
        }

        /// <summary>
        /// Every Outlook item which is to be synchronised must have a property SOModifiedDate, 
        /// a property SType, and a property SEntryId, referencing respectively the last time it
        /// was modified, the type of CRM item it is to be synchronised with, and the id of the
        /// CRM item it is to be synchronised with.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="modifiedDate">The value for the SOModifiedDate property.</param>
        /// <param name="type">The value for the SType property.</param>
        /// <param name="entryId">The value for the SEntryId property.</param>
        private static void EnsureSynchronisationPropertiesForOutlookItem(Outlook.AppointmentItem olItem, string modifiedDate, string type, string entryId)
        {
            EnsureSynchronisationPropertyForOutlookItem(olItem, "SOModifiedDate", modifiedDate);
            EnsureSynchronisationPropertyForOutlookItem(olItem, "SType", modifiedDate);
            EnsureSynchronisationPropertyForOutlookItem(olItem, "SEntryID", modifiedDate);
        }

        /// <summary>
        /// Ensure that this Outlook item has a property of this name with this value.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        private static void EnsureSynchronisationPropertyForOutlookItem(Outlook.AppointmentItem olItem, string name, string value)
        {
            Outlook.UserProperty olPropertyModifiedDate = olItem.UserProperties[name];
            if (olPropertyModifiedDate == null)
            {
                olPropertyModifiedDate = olItem.UserProperties.Add(name, Outlook.OlUserPropertyType.olText);
            }
            olPropertyModifiedDate.Value = value;
        }

        /// <summary>
        /// Add an item existing in CRM but not found in Outlook to Outlook.
        /// </summary>
        /// <param name="appointmentsFolder">The Outlook folder in which the item should be stored.</param>
        /// <param name="crmType">The CRM type of the item from which values are to be taken.</param>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <param name="date_start">The state date/time of the item, adjusted for timezone.</param>
        /// <returns></returns>
        private SyncState<Outlook.AppointmentItem> AddNewItemFromCrmToOutlook(Outlook.MAPIFolder appointmentsFolder, string crmType, dynamic crmItem, DateTime date_start)
        {
            Outlook.AppointmentItem aItem = appointmentsFolder.Items.Add(Outlook.OlItemType.olAppointmentItem);
            aItem.Subject = crmItem.name.value.ToString();
            aItem.Body = crmItem.description.value.ToString();
            if (!string.IsNullOrWhiteSpace(crmItem.date_start.value.ToString()))
            {
                aItem.Start = date_start;
                int iMin = 0, iHour = 0;
                if (!string.IsNullOrWhiteSpace(crmItem.duration_minutes.value.ToString()))
                {
                    iMin = int.Parse(crmItem.duration_minutes.value.ToString());
                }
                if (!string.IsNullOrWhiteSpace(crmItem.duration_hours.value.ToString()))
                {
                    iHour = int.Parse(crmItem.duration_hours.value.ToString());
                }
                if (crmType == "Meetings")
                {
                    aItem.Location = crmItem.location.value.ToString();
                    aItem.End = aItem.Start;
                    if (iHour > 0)
                        aItem.End.AddHours(iHour);
                    if (iMin > 0)
                        aItem.End.AddMinutes(iMin);
                }
                Log.Warn("   default SetRecepients");
                SetRecepients(aItem, crmItem.id.value.ToString(), crmType);

                try
                {
                    aItem.Duration = iMin + iHour * 60;
                }
                catch (Exception)
                {
                }
            }

            EnsureSynchronisationPropertiesForOutlookItem(aItem, crmItem.date_modified.value.ToString(), crmType, crmItem.id.value.ToString());

            var newState = new AppointmentSyncState(crmType)
            {
                OutlookItem = aItem,
                OModifiedDate = DateTime.ParseExact(crmItem.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null),
                CrmEntryId = crmItem.id.value.ToString(),
            };
            ItemsSyncState.Add(newState);
            aItem.Save();
            return newState;
        }

        /// <summary>
        /// Synchronise appointments in the specified folder with the specified SuiteCRM module.
        /// </summary>
        /// <param name="appointmentsFolder">The folder.</param>
        /// <param name="sModule">The module.</param>
        private void SyncFolder(Outlook.MAPIFolder appointmentsFolder, string sModule)
        {
            Log.Warn("SyncMeetings");
            try
            {
                /* this.ItemsSyncState already contains items to be synced */
                var untouched = new HashSet<SyncState<Outlook.AppointmentItem>>(this.ItemsSyncState);
                int nextOffset = -1; // offset of the next page of entries, if any.

                for (int iOffset = 0; nextOffset != 0; iOffset = nextOffset)
                {
                    /* get candidates for syncrhonisation from SuiteCRM one page at a time */
                    eGetEntryListResult entriesPage = clsSuiteCRMHelper.GetEntryList(sModule,
                        String.Format("assigned_user_id = '{0}'", clsSuiteCRMHelper.GetUserId()),
                        0, "date_start DESC", iOffset, false, 
                        clsSuiteCRMHelper.GetSugarFields(sModule));

                    nextOffset = entriesPage.next_offset; // get the offset of the next page

                    if (iOffset != nextOffset)
                    {
                        /* if there is a new page of candidates, add those candidates which were 
                         * not found in Outlook(?) to the untouched list */ 
                        UpdateAppointmentsFromCrmToOutlook(entriesPage.entry_list, appointmentsFolder, sModule, untouched);
                    }
                }

                eEntryValue[] invited = clsSuiteCRMHelper.getRelationships("Users", 
                    clsSuiteCRMHelper.GetUserId(), sModule.ToLower(), 
                    clsSuiteCRMHelper.GetSugarFields(sModule));
                if (invited != null)
                {
                    /* (?)likewise add those invitees not found in Outlook to the untouched list(?) */
                    UpdateAppointmentsFromCrmToOutlook(invited, appointmentsFolder, sModule, untouched);
                }

                try
                {
                    // TODO: unclear why this is only for 'meetings' and not for 'calls'. Bug #7?
                    if (sModule == "Meetings")
                    {
                        var itemsToBeDeletedFromOutlook = untouched.Where(a => a.ExistedInCrm && a.CrmType == sModule);
                        foreach (var item in itemsToBeDeletedFromOutlook)
                        {
                            try
                            {
                                item.OutlookItem.Delete();
                            }
                            catch (Exception)
                            {
                                Log.Warn("   Exception  oItem.oItem.Delete");
                            }
                            ItemsSyncState.Remove(item);
                        }
                    }

                    var itemsToBeAddedToCrm = untouched.Where(a => a.ShouldSyncWithCrm && !a.ExistedInCrm && a.CrmType == sModule);
                    foreach (var item in itemsToBeAddedToCrm)
                    {
                        AddItemFromOutlookToCrm(item.OutlookItem, sModule);
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

        /// <summary>
        /// Add this Outlook item, which may not exist in CRM, to CRM.
        /// </summary>
        /// <param name="olItem">The outlook item to add.</param>
        /// <param name="crmType">The CRM type to which it should be added</param>
        /// <param name="entryId">The id of this item in CRM, if known.</param>
        private void AddItemFromOutlookToCrm(Outlook.AppointmentItem olItem, string crmType, string entryId = "")
        {
            Log.Warn("AddItemFromOutlookToCrm");
            if (!settings.SyncCalendar)
                return;
            if (olItem != null)
            {
                try
                {
                    eNameValue[] data = new eNameValue[8];
                    DateTime uTCDateTime = new DateTime();
                    DateTime time2 = new DateTime();
                    uTCDateTime = olItem.Start.ToUniversalTime();
                    time2 = olItem.End.ToUniversalTime();
                    string str = string.Format("{0:yyyy-MM-dd HH:mm:ss}", uTCDateTime);
                    string str2 = string.Format("{0:yyyy-MM-dd HH:mm:ss}", time2);
                    int num = olItem.Duration / 60;
                    int num2 = olItem.Duration % 60;
                    data[0] = clsSuiteCRMHelper.SetNameValuePair("name", olItem.Subject);
                    data[1] = clsSuiteCRMHelper.SetNameValuePair("description", olItem.Body);
                    data[2] = clsSuiteCRMHelper.SetNameValuePair("location", olItem.Location);
                    data[3] = clsSuiteCRMHelper.SetNameValuePair("date_start", str);
                    data[4] = clsSuiteCRMHelper.SetNameValuePair("date_end", str2);
                    data[5] = clsSuiteCRMHelper.SetNameValuePair("duration_minutes", num2.ToString());
                    data[6] = clsSuiteCRMHelper.SetNameValuePair("duration_hours", num.ToString());

                    data[7] = String.IsNullOrEmpty(entryId) ?
                        clsSuiteCRMHelper.SetNameValuePair("assigned_user_id", clsSuiteCRMHelper.GetUserId()) :
                        clsSuiteCRMHelper.SetNameValuePair("id", entryId);

                    /* The id of the newly created CRM item */
                    string meetingId = clsSuiteCRMHelper.SetEntryUnsafe(data, crmType);
                    if (String.IsNullOrEmpty(entryId))
                    {
                        Log.Warn("    -- AddAppointmentToS AddAppointmentToS sID =" + entryId);

                        eSetRelationshipValue info = new eSetRelationshipValue
                        {
                            module2 = "meetings",
                            module2_id = meetingId,
                            module1 = "Users",
                            module1_id = clsSuiteCRMHelper.GetUserId()
                        };
                        clsSuiteCRMHelper.SetRelationshipUnsafe(info);

                    }
                    if (olItem.Recipients != null)
                    {
                        AddMeetingRecipientsFromOutlookToCrm(olItem, meetingId);
                    }

                    EnsureSynchronisationPropertiesForOutlookItem(olItem, DateTime.UtcNow.ToString(), crmType, meetingId);

                    Log.Warn("    AddItemFromOutlookToCrm Save ");
                    olItem.Save();

                    var sItem = this.ItemsSyncState.FirstOrDefault(a => a.OutlookItem.EntryID == olItem.EntryID);
                    if (sItem != null)
                    {
                        sItem.OutlookItem = olItem;
                        sItem.OModifiedDate = DateTime.UtcNow;
                        sItem.CrmEntryId = meetingId;
                        Log.Warn("    AddItemFromOutlookToCrm Edit ");
                    }
                    else
                    {
                        this.ItemsSyncState.Add(new AppointmentSyncState(crmType) { CrmEntryId = meetingId, OModifiedDate = DateTime.UtcNow, OutlookItem = olItem });
                        Log.Warn("    AddItemFromOutlookToCrm New ");
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("AppointementSyncing.AddItemFromOutlookToCrm", ex);
                }
            }
        }

        private void AddMeetingRecipientsFromOutlookToCrm(Outlook.AppointmentItem aItem, string meetingId)
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

                string sCID = SetCrmRelationshipFromOutlook(meetingId, objRecepient, "Contacts");
                if (sCID != "")
                {
                    string AccountID = clsSuiteCRMHelper.getRelationship("Contacts", sCID, "accounts");

                    if (AccountID != "")
                    {
                        eSetRelationshipValue info = new eSetRelationshipValue
                        {
                            module2 = "meetings",
                            module2_id = meetingId,
                            module1 = "Accounts",
                            module1_id = AccountID
                        };
                        clsSuiteCRMHelper.SetRelationshipUnsafe(info);
                    }
                    continue;
                }
                if (!String.IsNullOrEmpty(SetCrmRelationshipFromOutlook(meetingId, objRecepient, "Users"))) continue;
                if (!String.IsNullOrEmpty(SetCrmRelationshipFromOutlook(meetingId, objRecepient, "Leads"))) continue;
            }
        }

        /// <summary>
        /// Sets up a CRM relationship to mimic an Outlook relationship
        /// </summary>
        /// <param name="_result"></param>
        /// <param name="objRecepient"></param>
        /// <param name="relnName"></param>
        /// <returns></returns>
        private string SetCrmRelationshipFromOutlook(string _result, Outlook.Recipient objRecepient, string relnName)
        {
            string sCID = GetID(objRecepient.Address, relnName);
            if (sCID != "")
            {
                eSetRelationshipValue info = new eSetRelationshipValue
                {
                    module2 = "meetings",
                    module2_id = _result,
                    module1 = relnName,
                    module1_id = sCID
                };
                clsSuiteCRMHelper.SetRelationshipUnsafe(info);
            }

            return sCID;
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

        override public Outlook.MAPIFolder GetDefaultFolder()
        {
            return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
        }

        protected override bool IsCurrentView => Context.CurrentFolderItemType == Outlook.OlItemType.olAppointmentItem;

        // Should presumably be removed at some point. Existing code was ignoring deletions for Contacts and Tasks
        // (but not for Appointments).
        protected override bool PropagatesLocalDeletions => true;
    }
}
