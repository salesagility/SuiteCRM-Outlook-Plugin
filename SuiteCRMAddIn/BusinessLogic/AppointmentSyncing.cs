/**
 * Outlook integration for SuiteCRM.
 * @package Outlook integration for SuiteCRM
 * @copyright SalesAgility Ltd http://www.salesagility.com
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU LESSER GENERAL PUBLIC LICENCE as published by
 * the Free Software Foundation; either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU LESSER GENERAL PUBLIC LICENCE
 * along with this program; if not, see http://www.gnu.org/licenses
 * or write to the Free Software Foundation,Inc., 51 Franklin Street,
 * Fifth Floor, Boston, MA 02110-1301  USA
 *
 * @author SalesAgility <info@salesagility.com>
 */

namespace SuiteCRMAddIn.BusinessLogic
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using Newtonsoft.Json;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System.Text;

    /// <summary>
    /// Handles the synchronisation of appointments between Outlook and CMS.
    /// </summary>
    public class AppointmentSyncing: Synchroniser<Outlook.AppointmentItem>
    {
        /// <summary>
        /// The (primary) module I synchronise with.
        /// </summary>
        public const string CrmModule = "Meetings";

        /// <summary>
        /// The (other) module I synchronise with.
        /// </summary>
        /// <remarks>
        /// This rather makes me thing that there should be two classes here,
        /// CallsSynchroniser and MeetingsSynchroniser
        /// </remarks>
        public const string AltCrmModule = "Calls";

        public AppointmentSyncing(string name, SyncContext context)
            : base(name, context)
        {
            this.fetchQueryPrefix = "assigned_user_id = '{0}'";
        }

        public string GetID(string sEmailID, string sModule)
        {
            string str5 = "(" + sModule.ToLower() + ".id in (select eabr.bean_id from email_addr_bean_rel eabr INNER JOIN email_addresses ea on eabr.email_address_id = ea.id where eabr.bean_module = '" + sModule + "' and ea.email_address LIKE '%" + SuiteCRMAddIn.clsGlobals.MySqlEscape(sEmailID) + "%'))";

            Log.Info("-------------------GetID-----Start---------------");

            Log.Info("\tstr5=" + str5);

            Log.Info("-------------------GetID-----End---------------");

            string[] fields = new string[1];
            fields[0] = "id";
            eGetEntryListResult _result = clsSuiteCRMHelper.GetEntryList(sModule, str5, settings.SyncMaxRecords, "date_entered DESC", 0, false, fields);
            if (_result.result_count > 0)
            {
                return clsSuiteCRMHelper.GetValueByKey(_result.entry_list[0], "id");
            }
            return String.Empty;
        }

        override public Outlook.MAPIFolder GetDefaultFolder()
        {
            return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
        }

        public override bool SyncingEnabled => settings.SyncCalendar;

        protected override bool IsCurrentView => Context.CurrentFolderItemType == Outlook.OlItemType.olAppointmentItem;

        /// <summary>
        /// Entry point from event handler when an item is added in Outlook.
        /// </summary>
        /// <param name="appointment"></param>
        override protected void OutlookItemAdded(Outlook.AppointmentItem appointment)
        {
            LogItemAction(appointment, "AppointmentSyncing.OutlookItemAdded");

            if (IsCurrentView && !this.ItemsSyncState.Exists(a => a.OutlookItem.EntryID == appointment.EntryID))
            {
                AddOrUpdateItemFromOutlookToCrm(appointment, AppointmentSyncing.CrmModule);
            }
            else
            {
                Log.Warn(String.Format("AppointmentSyncing.OutlookItemAdded: item {0} had already been added", appointment.EntryID));
            }
        }

        protected override void SaveItem(Outlook.AppointmentItem olItem)
        {
            olItem.Save();
        }


        /// <summary>
        /// Entry point from event handler, called when an Outlook item of class AppointmentItem 
        /// has changed.
        /// </summary>
        /// <param name="olItem">The item which has changed.</param>
        override protected void OutlookItemChanged(Outlook.AppointmentItem olItem)
        {
            LogItemAction(olItem, "AppointmentSyncing.OutlookItemChanged");
            string entryId = olItem.EntryID;
            var syncStateForItem = GetSyncStateForItem(olItem);
            if (syncStateForItem != null)
            {
                var utcNow = DateTime.UtcNow;
                if (Math.Abs((utcNow - syncStateForItem.OModifiedDate).TotalSeconds) > 5.0)
                {
                    Log.Info("2 callitem.IsUpdate = " + syncStateForItem.IsUpdate);
                    syncStateForItem.IsUpdate = 0;
                }

                Log.Info("Before UtcNow - callitem.OModifiedDate= " + (int)(utcNow - syncStateForItem.OModifiedDate).TotalSeconds);

                if (Math.Abs((utcNow - syncStateForItem.OModifiedDate).TotalSeconds) > 2.0 && syncStateForItem.IsUpdate == 0)
                {
                    syncStateForItem.OModifiedDate = DateTime.UtcNow;
                    Log.Info("1 callitem.IsUpdate = " + syncStateForItem.IsUpdate);
                    syncStateForItem.IsUpdate++;
                }

                this.LogItemAction(syncStateForItem.OutlookItem, "AppointmentSyncState.OutlookItemChanged, syncing");
                Log.Info("utcNow= " + DateTime.UtcNow.ToString());
                Log.Info("UtcNow - callitem.OModifiedDate= " + (int)(DateTime.UtcNow - syncStateForItem.OModifiedDate).TotalSeconds);

                if (IsCurrentView && syncStateForItem.IsUpdate == 1)
                {
                    Outlook.UserProperty olPropertyType = olItem.UserProperties["SType"];
                    Outlook.UserProperty olPropertyEntryId = olItem.UserProperties["SEntryID"];
                    if (olPropertyType != null && olPropertyEntryId != null)
                    {
                        syncStateForItem.IsUpdate++;
                        AddOrUpdateItemFromOutlookToCrm(olItem, olPropertyType.Value.ToString(), olPropertyEntryId.Value.ToString());
                    }
                }
            }
            else
            {
                /* we don't have a sync state for this item (presumably formerly private);
                 *  that's OK, treat it as new */
                OutlookItemAdded(olItem);
            }
        }

        // Should presumably be removed at some point. Existing code was ignoring deletions for Contacts and Tasks
        // (but not for Appointments).
        protected override bool PropagatesLocalDeletions => true;

        /// <summary>
        /// Action method of the thread.
        /// </summary>
        public override void SynchroniseAll()
        {
            base.SynchroniseAll();
            SyncFolder(GetDefaultFolder(), AppointmentSyncing.AltCrmModule);
        }

        public override string DefaultCrmModule
        {
            get
            {
                return AppointmentSyncing.CrmModule;
            }
        }

        /// <summary>
        /// Ensure that this Outlook item has a property of this name with this value.
        /// </summary>
        /// <remarks>
        /// TODO: Candidate for refactoring to superclass.
        /// </remarks>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        protected override void EnsureSynchronisationPropertyForOutlookItem(Outlook.AppointmentItem olItem, string name, string value)
        {
            Outlook.UserProperty olProperty = olItem.UserProperties[name];
            if (olProperty == null)
            {
                olProperty = olItem.UserProperties.Add(name, Outlook.OlUserPropertyType.olText);
            }
            olProperty.Value = value;
        }


        private void AddCurrentUserAsOwner(Outlook.AppointmentItem olItem, string meetingId)
        {
            LogItemAction(olItem, "AppointmentSyncing.AddItemFromOutlookToCrm, adding current user");

            eSetRelationshipValue info = new eSetRelationshipValue
            {
                module2 = AppointmentSyncing.CrmModule,
                module2_id = meetingId,
                module1 = "Users",
                module1_id = clsSuiteCRMHelper.GetUserId()
            };
            clsSuiteCRMHelper.SetRelationshipUnsafe(info);
        }

        private void AddMeetingRecipientsFromOutlookToCrm(Outlook.AppointmentItem aItem, string meetingId)
        {
            LogItemAction(aItem, "AppointmentSyncing.AddMeetingRecipientsFromOutlookToCrm");
            foreach (Outlook.Recipient objRecepient in aItem.Recipients)
            {
                try
                {
                    Log.Info("objRecepientName= " + objRecepient.Name.ToString());
                    Log.Info("objRecepient= " + objRecepient.Address.ToString());
                }
                catch
                {
                    Log.Warn("objRecepient ERROR");
                    continue;
                }

                string sCID = SetCrmRelationshipFromOutlook(meetingId, objRecepient, ContactSyncing.CrmModule);
                if (sCID != String.Empty)
                {
                    string AccountID = clsSuiteCRMHelper.getRelationship(ContactSyncing.CrmModule, sCID, "accounts");

                    if (AccountID != String.Empty)
                    {
                        eSetRelationshipValue info = new eSetRelationshipValue
                        {
                            module2 = AppointmentSyncing.CrmModule,
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
        /// Add an item existing in CRM but not found in Outlook to Outlook.
        /// </summary>
        /// <param name="appointmentsFolder">The Outlook folder in which the item should be stored.</param>
        /// <param name="crmType">The CRM type of the item from which values are to be taken.</param>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <param name="date_start">The state date/time of the item, adjusted for timezone.</param>
        /// <returns>A sync state object for the new item.</returns>
        private SyncState<Outlook.AppointmentItem> AddNewItemFromCrmToOutlook(
            Outlook.MAPIFolder appointmentsFolder,
            string crmType,
            eEntryValue crmItem,
            DateTime date_start)
        {
            Outlook.AppointmentItem olItem = appointmentsFolder.Items.Add(Outlook.OlItemType.olAppointmentItem);
            olItem.Subject = crmItem.GetValueAsString("name");
            olItem.Body = crmItem.GetValueAsString("description");

            LogItemAction(olItem, "AppointmentSyncing.AddNewItemFromCrmToOutlook");

            if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("date_start")))
            {
                olItem.Start = date_start;
                int iMin = 0, iHour = 0;
                if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("duration_minutes")))
                {
                    iMin = int.Parse(crmItem.GetValueAsString("duration_minutes"));
                }
                if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("duration_hours")))
                {
                    iHour = int.Parse(crmItem.GetValueAsString("duration_hours"));
                }
                if (crmType == AppointmentSyncing.CrmModule)
                {
                    olItem.Location = crmItem.GetValueAsString("location");
                    olItem.End = olItem.Start;
                    if (iHour > 0)
                        olItem.End.AddHours(iHour);
                    if (iMin > 0)
                        olItem.End.AddMinutes(iMin);
                }
                Log.Info("\tdefault SetRecepients");
                SetRecipients(olItem, crmItem.GetValueAsString("id"), crmType);

                try
                {
                    olItem.Duration = iMin + iHour * 60;
                }
                catch (Exception)
                {
                }
            }

            string crmId = crmItem.GetValueAsString("id");
            EnsureSynchronisationPropertiesForOutlookItem(olItem, crmItem.GetValueAsString("date_modified"), crmType, crmId);

            var newState = new AppointmentSyncState(crmType)
            {
                OutlookItem = olItem,
                OModifiedDate = DateTime.ParseExact(crmItem.GetValueAsString("date_modified"), "yyyy-MM-dd HH:mm:ss", null),
                CrmEntryId = crmId,
            };
            ItemsSyncState.Add(newState);
            olItem.Save();

            LogItemAction(newState.OutlookItem, "AppointmentSyncing.AddNewItemFromCrmToOutlook, saved item");

            return newState;
        }

        /// <summary>
        /// Add this Outlook item, which may not exist in CRM, to CRM.
        /// </summary>
        /// <param name="olItem">The outlook item to add.</param>
        /// <param name="crmType">The CRM type to which it should be added</param>
        /// <param name="entryId">The id of this item in CRM, if known (in which case I should be doing
        /// an update, not an add).</param>
        protected override string AddOrUpdateItemFromOutlookToCrm(Outlook.AppointmentItem olItem, string crmType, string entryId = "")
        {
            string result = entryId;

            if (SyncingEnabled && olItem != null)
            {
                if (ShouldDeleteFromCrm(olItem))
                {
                    /* Issue #14: if it is non-public, it should be removed from or not copied to CRM */
                    LogItemAction(olItem, "AppointmentSyncing.AddItemFromOutlookToCrm Deleting");
                    var syncStateForItem = this.GetSyncStateForItem(olItem);

                    DeleteFromCrm(olItem);
                }
                else if (ShouldDespatchToCrm(olItem))
                {
                    result = base.AddOrUpdateItemFromOutlookToCrm(olItem, crmType, entryId);

                    if (String.IsNullOrEmpty(entryId))
                    {
                        AddCurrentUserAsOwner(olItem, result);
                    }
                    if (olItem.Recipients != null)
                    {
                        AddMeetingRecipientsFromOutlookToCrm(olItem, result);
                    }
                }
                else
                {
                    LogItemAction(olItem, "AppointmentSyncing.AddItemFromOutlookToCrm, Not despatching");
                }
            }

            return result;
        }

        /// <summary>
        /// Construct a JSON packet representing this Outlook item, and despatch it to CRM. 
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmType">The type within CRM to which the item should be added.</param>
        /// <param name="entryId">??</param>
        /// <returns>The CRM id of the object created or modified.</returns>
        protected override string ConstructAndDespatchCrmItem(Outlook.AppointmentItem olItem, string crmType, string entryId)
        {
            List<eNameValue> data = new List<eNameValue>();

            DateTime uTCDateTime = new DateTime();
            DateTime time2 = new DateTime();
            uTCDateTime = olItem.Start.ToUniversalTime();
            time2 = olItem.End.ToUniversalTime();
            string str = string.Format("{0:yyyy-MM-dd HH:mm:ss}", uTCDateTime);
            string str2 = string.Format("{0:yyyy-MM-dd HH:mm:ss}", time2);
            int num = olItem.Duration / 60;
            int num2 = olItem.Duration % 60;
            data.Add(clsSuiteCRMHelper.SetNameValuePair("name", olItem.Subject));
            data.Add(clsSuiteCRMHelper.SetNameValuePair("description", olItem.Body));
            data.Add(clsSuiteCRMHelper.SetNameValuePair("location", olItem.Location));
            data.Add(clsSuiteCRMHelper.SetNameValuePair("date_start", str));
            data.Add(clsSuiteCRMHelper.SetNameValuePair("date_end", str2));
            data.Add(clsSuiteCRMHelper.SetNameValuePair("duration_minutes", num2.ToString()));
            data.Add(clsSuiteCRMHelper.SetNameValuePair("duration_hours", num.ToString()));

            data.Add(String.IsNullOrEmpty(entryId) ?
                clsSuiteCRMHelper.SetNameValuePair("assigned_user_id", clsSuiteCRMHelper.GetUserId()) :
                clsSuiteCRMHelper.SetNameValuePair("id", entryId));

            /* The id of the newly created or modified CRM item */
            return clsSuiteCRMHelper.SetEntryUnsafe(data.ToArray(), crmType);
        }

        /// <summary>
        /// Delete this Outlook item from CRM, and tidy up afterwards.
        /// </summary>
        /// <param name="olItem">The Outlook item to delete.</param>
        private void DeleteFromCrm(Outlook.AppointmentItem olItem)
        {
            if (olItem != null)
            {
                /* Remove the magic properties */
                RemoveSynchronisationPropertiesFromOutlookItem(olItem);
                SyncState<Outlook.AppointmentItem> syncStateForItem = GetSyncStateForItem(olItem);
                if (syncStateForItem != null)
                {
                    this.RemoveFromCrm(syncStateForItem);
                    /* TODO: Not at all sure I should remove the sync state */
                    RemoveItemSyncState(syncStateForItem);
                }
            }
        }

        /// <summary>
        /// Get all items in this appointments folder. Should be called just once (per folder?) 
        /// when the add-in starts up; initialises the SyncState list.  
        /// </summary>
        /// <param name="appointmentsFolder">The folder to scan.</param>
        protected override void GetOutlookItems(Outlook.MAPIFolder appointmentsFolder)
        {
            try
            {
                if (ItemsSyncState == null)
                {
                    ItemsSyncState = new ThreadSafeList<SyncState<Outlook.AppointmentItem>>();

                    foreach (Outlook.AppointmentItem aItem in appointmentsFolder.Items)
                    {
                        if (aItem.Start >= this.GetStartDate())
                        {
                            Outlook.UserProperty olPropertyModified = aItem.UserProperties["SOModifiedDate"];
                            if (olPropertyModified != null)
                            {
                                /* The appointment probably already has the three magic properties 
                                 * required for synchronisation; is that a proxy for believing that it
                                 * already exists in CRM? If so, is it reliable? */
                                Outlook.UserProperty olPropertyType = aItem.UserProperties["SType"];
                                Outlook.UserProperty olPropertyEntryId = aItem.UserProperties["SEntryID"];
                                var crmType = olPropertyType.Value.ToString();
                                ItemsSyncState.Add(new AppointmentSyncState(crmType)
                                {
                                    OutlookItem = aItem,
                                    OModifiedDate = DateTime.UtcNow,
                                    CrmEntryId = olPropertyEntryId.Value.ToString()
                                });
                                LogItemAction(aItem, "AppointmentSyncing.GetOutlookItems: Adding known item to queue");
                            }
                            else
                            {
                                ItemsSyncState.Add(new AppointmentSyncState(AppointmentSyncing.CrmModule)
                                {
                                    OutlookItem = aItem,
                                });
                                LogItemAction(aItem, "AppointmentSyncing.GetOutlookItems: Adding unknown item to queue");
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

        /// <summary>
        /// Get the unique sync state related to this Outlook item from among my sync states, if present.
        /// </summary>
        /// <param name="olItem">The Outlook item to seek</param>
        /// <returns>The sync state related to that item.</returns>
        private SyncState<Outlook.AppointmentItem> GetSyncStateForItem(Outlook.AppointmentItem olItem)
        {
            SyncState<Outlook.AppointmentItem> result;
            try
            {
                /* if there are duplicate entries I want them logged */
                result = this.ItemsSyncState.SingleOrDefault(a => a.OutlookItem.EntryID == olItem.EntryID);
            }
            catch (InvalidOperationException notUnique)
            {
                Log.Error(
                    String.Format(
                        "AppointmentSyncing.AddItemFromOutlookToCrm: Id {0} was not unique in this.ItemsSyncState?",
                        olItem.EntryID),
                    notUnique);

                result = this.ItemsSyncState.FirstOrDefault(a => a.OutlookItem.EntryID == olItem.EntryID);
            }

            return result;
        }

        /// <summary>
        /// Log a message regarding this Outlook appointment.
        /// </summary>
        /// <param name="olItem">The outlook item.</param>
        /// <param name="message">The message to be logged.</param>
        protected override void LogItemAction(Outlook.AppointmentItem olItem, string message)
        {
            try
            {
                Outlook.UserProperty olPropertyEntryId = olItem.UserProperties["SEntryID"];
                string crmId = olPropertyEntryId == null ?
                    "[not present]" :
                    olPropertyEntryId.Value;
                StringBuilder bob = new StringBuilder();
                bob.Append($"{message}:\n\tOutlook Id  : {olItem.EntryID}\n\tCRM Id      : {crmId}\n\tSubject     : '{olItem.Subject}'\n\tSensitivity : {olItem.Sensitivity}\n\tRecipients");
                foreach (var recipient in olItem.Recipients)
                {
                    bob.Append($"\t\t{recipient}\n");
                }
                Log.Info(bob.ToString());
            }
            catch (COMException)
            {
                // Ignore: happens if the outlook item is already deleted.
            }
        }

        /// <summary>
        /// Update a single appointment in the specified Outlook folder with changes from CRM, but 
        /// only if its start date is fewer than five days in the past.
        /// </summary>
        /// <param name="folder">The folder to synchronise into.</param>
        /// <param name="crmType">The CRM type of the candidate item.</param>
        /// <param name="candidateItem">The candidate item from CRM.</param>
        /// <returns>The synchronisation state of the item updated (if it was updated).</returns>
        protected override SyncState<Outlook.AppointmentItem> UpdateFromCrm(
            Outlook.MAPIFolder folder,
            string crmType,
            eEntryValue crmItem)
        {
            SyncState<Outlook.AppointmentItem> result = null;
            DateTime date_start = DateTime.ParseExact(crmItem.GetValueAsString("date_start"), "yyyy-MM-dd HH:mm:ss", null);
            date_start = date_start.Add(new DateTimeOffset(DateTime.Now).Offset); // correct for offset from UTC.
            if (date_start >= GetStartDate())
            {
                /* search for the item among the items I already know about */
                var oItem = this.ItemsSyncState.FirstOrDefault(a => a.CrmEntryId == crmItem.GetValueAsString("id") && a.CrmType == crmType);
                if (oItem == null)
                {
                    /* didn't find it, so add it to Outlook */
                    result = AddNewItemFromCrmToOutlook(folder, crmType, crmItem, date_start);
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
        /// Remove the synchronisation properties from this Outlook item.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        private static void RemoveSynchronisationPropertiesFromOutlookItem(Outlook.AppointmentItem olItem)
        {
            RemoveSynchronisationPropertyFromOutlookItem(olItem, "SEntryId");
            RemoveSynchronisationPropertyFromOutlookItem(olItem, "SType");
            RemoveSynchronisationPropertyFromOutlookItem(olItem, "SOModifiedDate");
        }

        /// <summary>
        /// Ensure that this Outlook item does not have a property of this name.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="name">The name.</param>
        private static void RemoveSynchronisationPropertyFromOutlookItem(Outlook.AppointmentItem olItem, string name)
        {
            int found = 0;
            /* typical Microsoft, you can only remove a user property by its 1-based number */
            for (int i = 1; i <= olItem.UserProperties.Count; i++)
            {
                if (olItem.UserProperties[i].Name == name)
                {
                    found = i;
                    break;
                }
            }

            if (found > 0)
            {
                olItem.UserProperties.Remove(found);
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
            if (sCID != String.Empty)
            {
                eSetRelationshipValue info = new eSetRelationshipValue
                {
                    module2 = AppointmentSyncing.CrmModule,
                    module2_id = _result,
                    module1 = relnName,
                    module1_id = sCID
                };
                clsSuiteCRMHelper.SetRelationshipUnsafe(info);
            }

            return sCID;
        }

        private void SetRecipients(Outlook.AppointmentItem olAppointment, string sMeetingID, string sModule)
        {
            olAppointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
            int iCount = olAppointment.Recipients.Count;
            for (int iItr = 1; iItr <= iCount; iItr++)
            {
                olAppointment.Recipients.Remove(1);
            }

            string[] invitee_categories = { "users", ContactSyncing.CrmModule, "leads" };
            foreach (string invitee_category in invitee_categories)
            {
                eEntryValue[] relationships = clsSuiteCRMHelper.getRelationships(sModule, sMeetingID, invitee_category, new string[] { "id", "email1", "phone_work" });
                if (relationships != null)
                {

                    foreach (var relationship in relationships)
                    {
                        string phone_work = relationship.GetValueAsString("phone_work");
                        string sTemp =
                            (sModule == AppointmentSyncing.CrmModule) || String.IsNullOrWhiteSpace(phone_work) ?
                                relationship.GetValueAsString("email1") :
                                relationship.GetValueAsString("email1") + ":" + phone_work;

                        if (!String.IsNullOrWhiteSpace(sTemp))
                        {
                            olAppointment.Recipients.Add(sTemp);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// We should delete an item from CRM if it already exists in CRM, but it is now private.
        /// </summary>
        /// <param name="olItem">The Outlook item</param>
        /// <returns>true if the Outlook item should be deleted from CRM.</returns>
        private bool ShouldDeleteFromCrm(Outlook.AppointmentItem olItem)
        {
            Outlook.UserProperty olPropertyEntryId = olItem.UserProperties["SEntryID"];
            bool result = (olPropertyEntryId != null && olItem.Sensitivity != Outlook.OlSensitivity.olNormal);

            LogItemAction(olItem, $"ShouldDeleteFromCrm returning {result}");

            return result;
        }

        /// <summary>
        /// True if we should despatch this item to CRM, else false.
        /// </summary>
        /// <param name="olItem"></param>
        /// <returns>true iff settings.SyncCalendar is true, the item is not null, and it is not private (normal sensitivity)</returns>
        private bool ShouldDespatchToCrm(Outlook.AppointmentItem olItem)
        {
            return olItem != null && settings.SyncCalendar && olItem.Sensitivity == Outlook.OlSensitivity.olNormal;
        }

        /// <summary>
        /// Synchronise items in the specified folder with the specified SuiteCRM module.
        /// </summary>
        /// <remarks>
        /// TODO: candidate for refactoring upwards, in concert with ContactSyncing.SyncFolder.
        /// </remarks>
        /// <param name="folder">The folder.</param>
        /// <param name="crmModule">The module.</param>
        protected override void SyncFolder(Outlook.MAPIFolder folder, string crmModule)
        {
            Log.Info(String.Format("AppointmentSyncing.SyncFolder: '{0}'", crmModule));
            try
            {
                /* this.ItemsSyncState already contains items to be synced. */
                var untouched = new HashSet<SyncState<Outlook.AppointmentItem>>(this.ItemsSyncState);
                MergeRecordsFromCrm(folder, crmModule, untouched);

                eEntryValue[] invited = clsSuiteCRMHelper.getRelationships("Users",
                    clsSuiteCRMHelper.GetUserId(), crmModule.ToLower(),
                    clsSuiteCRMHelper.GetSugarFields(crmModule));
                if (invited != null)
                {
                    UpdateItemsFromCrmToOutlook(invited, folder, untouched, crmModule);
                }

                try
                {
                    this.ResolveUnmatchedItems(untouched, crmModule);
                }
                catch (Exception ex)
                {
                    Log.Error("AppointmentSyncing.SyncFolder: Exception", ex);
                }
            }
            catch (Exception ex)
            {
                Log.Error("AppointmentSyncing.SyncFolder: Exception", ex);
            }
        }

        /// <summary>
        /// Update an existing Outlook item with values taken from a corresponding CRM item. Note that 
        /// this just overwrites all values in the Outlook item.
        /// </summary>
        /// <param name="crmType">The CRM type of the item from which values are to be taken.</param>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <param name="date_start">The state date/time of the item, adjusted for timezone.</param>
        /// <param name="oItem">The outlook item assumed to correspond with the CRM item.</param>
        /// <returns>An appropriate sync state.</returns>
        private SyncState<Outlook.AppointmentItem> UpdateExistingOutlookItemFromCrm(
            string crmType, 
            eEntryValue crmItem, 
            DateTime date_start, 
            SyncState<Outlook.AppointmentItem> oItem)
        {
            LogItemAction(oItem.OutlookItem, "AppointmentSyncing.UpdateExistingOutlookItemFromCrm");
            Outlook.AppointmentItem olAppointment = oItem.OutlookItem;
            Outlook.UserProperty olPropertyModifiedDate = olAppointment.UserProperties["SOModifiedDate"];

            if (olPropertyModifiedDate.Value != crmItem.GetValueAsString("date_modified"))
            {
                olAppointment.Subject = crmItem.GetValueAsString("name");
                olAppointment.Body = crmItem.GetValueAsString("description");
                if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("date_start")))
                {
                    UpdateOutlookStartAndDuration(crmType, crmItem, date_start, olAppointment);
                }

                EnsureSynchronisationPropertiesForOutlookItem(olAppointment, crmItem.GetValueAsString("date_modified"), crmType, crmItem.id);
                olAppointment.Save();
                LogItemAction(oItem.OutlookItem, "AppointmentSyncing.UpdateExistingOutlookItemFromCrm, item saved");
            }
            Log.Warn((string)("Not default dResult.date_modified= " + crmItem.GetValueAsString("date_modified")));
            oItem.OModifiedDate = DateTime.ParseExact(crmItem.GetValueAsString("date_modified"), "yyyy-MM-dd HH:mm:ss", null);

            return oItem;
        }

        /// <summary>
        /// Update this Outlook appointment's start and duration from this CRM object.
        /// </summary>
        /// <param name="crmType">The CRM type of the item from which values are to be taken.</param>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <param name="date_start">The state date/time of the item, adjusted for timezone.</param>
        /// <param name="olAppointment">The outlook item assumed to correspond with the CRM item.</param>
        private void UpdateOutlookStartAndDuration(string crmType, eEntryValue crmItem, DateTime date_start, Outlook.AppointmentItem olAppointment)
        {
            olAppointment.Start = date_start;
            var minutesString = crmItem.GetValueAsString("duration_minutes");
            var hoursString = crmItem.GetValueAsString("duration_hours");

            int minutes = string.IsNullOrWhiteSpace(minutesString) ? 0 : int.Parse(minutesString);
            int hours = string.IsNullOrWhiteSpace(hoursString) ? 0 : int.Parse(hoursString);

            if (crmType == AppointmentSyncing.CrmModule)
            {
                olAppointment.Location = crmItem.GetValueAsString("location");
                olAppointment.End = olAppointment.Start;
                if (hours > 0)
                    olAppointment.End.AddHours(hours);
                if (minutes > 0)
                    olAppointment.End.AddMinutes(minutes);
                Log.Info("\tSetRecepients");
                SetRecipients(olAppointment, crmItem.GetValueAsString("id"), crmType);
            }
            olAppointment.Duration = minutes + hours * 60;
        }

        protected override SyncState<Outlook.AppointmentItem> ConstructSyncState(Outlook.AppointmentItem oItem)
        {
            return new AppointmentSyncState(oItem.UserProperties["SType"]?.Value.ToString())
            {
                OutlookItem = oItem,
                CrmEntryId = oItem.UserProperties["SEntryID"]?.Value.ToString(),
                OModifiedDate = ParseDateTimeFromUserProperty(oItem.UserProperties["SOModifiedDate"]?.Value.ToString()),
            };
        }

        protected override SyncState<Outlook.AppointmentItem> GetExistingSyncState(Outlook.AppointmentItem oItem)
        {
            return ItemsSyncState.FirstOrDefault(a => a.OutlookItem.EntryID == oItem.EntryID);
        }
    }
}
