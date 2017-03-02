﻿/**
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

    /// <summary>
    /// Handles the synchronisation of appointments between Outlook and CMS.
    /// </summary>
    public class AppointmentSyncing: Synchroniser<Outlook.AppointmentItem>
    {
        public AppointmentSyncing(string name, SyncContext context)
            : base(name, context)
        {
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
                AddOrUpdateItemFromOutlookToCrm(appointment, "Meetings");
            }
            else
            {
                Log.Warn(String.Format("AppointmentSyncing.OutlookItemAdded: item {0} had already been added", appointment.EntryID));
            }
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
            AddSuiteCrmOutlookCategory();
            Outlook.MAPIFolder folder = GetDefaultFolder();

            GetOutlookItems(folder);
            SyncFolder(folder, "Meetings");
            SyncFolder(folder, "Calls");
        }

        private void AddCurrentUserAsOwner(Outlook.AppointmentItem olItem, string meetingId)
        {
            LogItemAction(olItem, "AppointmentSyncing.AddItemFromOutlookToCrm, adding current user");

            eSetRelationshipValue info = new eSetRelationshipValue
            {
                module2 = "meetings",
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

                string sCID = SetCrmRelationshipFromOutlook(meetingId, objRecepient, "Contacts");
                if (sCID != String.Empty)
                {
                    string AccountID = clsSuiteCRMHelper.getRelationship("Contacts", sCID, "accounts");

                    if (AccountID != String.Empty)
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
            dynamic crmItem,
            DateTime date_start)
        {
            Outlook.AppointmentItem olItem = appointmentsFolder.Items.Add(Outlook.OlItemType.olAppointmentItem);

            LogItemAction(olItem, "AppointmentSyncing.AddNewItemFromCrmToOutlook");

            olItem.Subject = crmItem.name.value.ToString();
            olItem.Body = crmItem.description.value.ToString();
            if (!string.IsNullOrWhiteSpace(crmItem.date_start.value.ToString()))
            {
                olItem.Start = date_start;
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
                    olItem.Location = crmItem.location.value.ToString();
                    olItem.End = olItem.Start;
                    if (iHour > 0)
                        olItem.End.AddHours(iHour);
                    if (iMin > 0)
                        olItem.End.AddMinutes(iMin);
                }
                Log.Info("\tdefault SetRecepients");
                SetRecipients(olItem, crmItem.id.value.ToString(), crmType);

                try
                {
                    olItem.Duration = iMin + iHour * 60;
                }
                catch (Exception)
                {
                }
            }

            string crmId = crmItem.id.value.ToString();
            EnsureSynchronisationPropertiesForOutlookItem(olItem, crmItem.date_modified.value.ToString(), crmType, crmId);

            var newState = new AppointmentSyncState(crmType)
            {
                OutlookItem = olItem,
                OModifiedDate = DateTime.ParseExact(crmItem.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null),
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
        private void AddOrUpdateItemFromOutlookToCrm(Outlook.AppointmentItem olItem, string crmType, string entryId = "")
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
                LogItemAction(olItem, "AppointmentSyncing.AddItemFromOutlookToCrm, Despatching");

                try
                {
                    string meetingId = ConstructAndDespatchCrmItem(olItem, crmType, entryId);

                    if (String.IsNullOrEmpty(entryId))
                    {
                        AddCurrentUserAsOwner(olItem, meetingId);
                    }
                    if (olItem.Recipients != null)
                    {
                        AddMeetingRecipientsFromOutlookToCrm(olItem, meetingId);
                    }

                    /* this is where the CRM entry id gets fixed up in Outlook */
                    EnsureSynchronisationPropertiesForOutlookItem(olItem, DateTime.UtcNow.ToString(), crmType, meetingId);

                    LogItemAction(olItem, "AppointmentSyncing.AddItemFromOutlookToCrm Save");
                    olItem.Save();

                    /* Find the existing syncstate whose outlook item has the same EntryId value as the current olItem */
                    var syncStateForItem = this.GetSyncStateForItem(olItem);

                    if (syncStateForItem != null)
                    {
                        syncStateForItem.OutlookItem = olItem;
                        syncStateForItem.OModifiedDate = DateTime.UtcNow;
                        syncStateForItem.CrmEntryId = meetingId;
                        LogItemAction(olItem, "AppointmentSyncing.AddItemFromOutlookToCrm Edit sync state");
                    }
                    else
                    {
                        this.ItemsSyncState.Add(new AppointmentSyncState(crmType) { CrmEntryId = meetingId, OModifiedDate = DateTime.UtcNow, OutlookItem = olItem });
                        LogItemAction(olItem, "AppointmentSyncing.AddItemFromOutlookToCrm Create sync state");
                    }

                }
                catch (Exception ex)
                {
                    Log.Error("AppointementSyncing.AddItemFromOutlookToCrm", ex);
                }
            }
            else
            {
                LogItemAction(olItem, "AppointmentSyncing.AddItemFromOutlookToCrm, Not despatching");
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

        /// <summary>
        /// Construct a JSON packet representing this Outlook item, and despatch it to CRM. 
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmType">The type within CRM to which the item should be added.</param>
        /// <param name="entryId">??</param>
        /// <returns>The CRM id of the object created or modified.</returns>
        private string ConstructAndDespatchCrmItem(Outlook.AppointmentItem olItem, string crmType, string entryId)
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

            /* The id of the newly created or modified CRM item */
            return clsSuiteCRMHelper.SetEntryUnsafe(data, crmType);
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
        /// Every Outlook item which is to be synchronised must have a property SOModifiedDate, 
        /// a property SType, and a property SEntryId, referencing respectively the last time it
        /// was modified, the type of CRM item it is to be synchronised with, and the id of the
        /// CRM item it is to be synchronised with.
        /// </summary>
        /// <remarks>
        /// TODO: Candidate for refactoring to superclass.
        /// </remarks>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="modifiedDate">The value for the SOModifiedDate property.</param>
        /// <param name="type">The value for the SType property.</param>
        /// <param name="entryId">The value for the SEntryId property.</param>
        private static void EnsureSynchronisationPropertiesForOutlookItem(Outlook.AppointmentItem olItem, string modifiedDate, string type, string entryId)
        {
            EnsureSynchronisationPropertyForOutlookItem(olItem, "SOModifiedDate", modifiedDate);
            EnsureSynchronisationPropertyForOutlookItem(olItem, "SType", type);
            EnsureSynchronisationPropertyForOutlookItem(olItem, "SEntryID", entryId);
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
        private static void EnsureSynchronisationPropertyForOutlookItem(Outlook.AppointmentItem olItem, string name, string value)
        {
            Outlook.UserProperty olProperty = olItem.UserProperties[name];
            if (olProperty == null)
            {
                olProperty = olItem.UserProperties.Add(name, Outlook.OlUserPropertyType.olText);
            }
            olProperty.Value = value;
        }

        /// <summary>
        /// Get all items in this appointments folder. Should be called just once (per folder?) 
        /// when the add-in starts up; initialises the SyncState list.  
        /// </summary>
        /// <param name="appointmentsFolder">The folder to scan.</param>
        private void GetOutlookItems(Outlook.MAPIFolder appointmentsFolder)
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
                                ItemsSyncState.Add(new AppointmentSyncState("Meetings")
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
        private void LogItemAction(Outlook.AppointmentItem olItem, string message)
        {
            try
            {
                Outlook.UserProperty olPropertyEntryId = olItem.UserProperties["SEntryID"];
                string crmId = olPropertyEntryId == null ?
                    "[not present]" :
                    olPropertyEntryId.Value;
                Log.Info(
                    String.Format("{0}:\n\tOutlook Id  : {1}\n\tCRM Id      : {2}\n\tSubject     : '{3}'\n\tSensitivity : {4}",
                    message, olItem.EntryID, crmId, olItem.Subject, olItem.Sensitivity));
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
        private SyncState<Outlook.AppointmentItem> MaybeUpdateAppointmentFromCrmToOutlook(
            Outlook.MAPIFolder folder,
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
        /// Remove an outlook item and its associated sync state.
        /// </summary>
        /// <remarks>
        /// TODO: candidate for refactoring to superclass.
        /// </remarks>
        /// <param name="syncState">The sync state of the item to remove.</param>
        private void RemoveItemAndSyncState(SyncState<Outlook.AppointmentItem> syncState)
        {
            this.LogItemAction(syncState.OutlookItem, "AppointmentSyncing.SyncFolder, deleting item");
            try
            {
                syncState.OutlookItem.Delete();
            }
            catch (Exception ex)
            {
                Log.Error("AppointmentSyncing.SyncFolder: Exception  oItem.oItem.Delete", ex);
            }
            this.RemoveItemSyncState(syncState);
        }

        /// <summary>
        /// Remove an item from ItemsSyncState.
        /// </summary>
        /// <remarks>
        /// TODO: candidate for refactoring to superclass.
        /// </remarks>
        /// <param name="item">The sync state of the item to remove.</param>
        private void RemoveItemSyncState(SyncState<Outlook.AppointmentItem> item)
        {
            this.LogItemAction(item.OutlookItem, "AppointmentSyncing.RemoveItemSyncState, removed item from queue");
            this.ItemsSyncState.Remove(item);
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
                    module2 = "meetings",
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

            string[] invitee_categories = { "users", "contacts", "leads" };
            foreach (string invitee_category in invitee_categories)
            {
                eEntryValue[] Users = clsSuiteCRMHelper.getRelationships(sModule, sMeetingID, invitee_category, new string[] { "id", "email1", "phone_work" });
                if (Users != null)
                {

                    foreach (var oResult1 in Users)
                    {
                        dynamic dResult1 = JsonConvert.DeserializeObject(oResult1.name_value_object.ToString());

                        Log.Info("-------------------SetRecepients-----Start-----dResult1---2-------");
                        Log.Info((string)Convert.ToString(dResult1));
                        Log.Info("-------------------SetRecepients-----End---------------");

                        string phone_work = dResult1.phone_work.value.ToString();
                        string sTemp =
                            (sModule == "Meetings") || String.IsNullOrWhiteSpace(phone_work) ?
                                dResult1.email1.value.ToString() :
                                dResult1.email1.value.ToString() + ":" + phone_work;

                        if (!String.IsNullOrWhiteSpace(sTemp))
                        {
                            olAppointment.Recipients.Add(sTemp);
                        }
                    }
                }
            }
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
        /// We should delete an item from CRM if it already exists in CRM, but it is now private.
        /// TODO: it should also be deleted from CRM if it's deleted from Outlook.
        /// </summary>
        /// <param name="olItem">The Outlook item</param>
        /// <returns>true if the Outlook item should be deleted from CRM.</returns>
        private bool ShouldDeleteFromCrm(Outlook.AppointmentItem olItem)
        {
            Outlook.UserProperty olPropertyEntryId = olItem.UserProperties["SEntryID"];
            bool result = (olPropertyEntryId != null && olItem.Sensitivity != Outlook.OlSensitivity.olNormal);

            LogItemAction(olItem, String.Format( "ShouldDeleteFromCrm returning {0}", result));

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
        private void SyncFolder(Outlook.MAPIFolder folder, string crmModule)
        {
            Log.Info(String.Format("AppointmentSyncing.SyncFolder: '{0}'", crmModule));
            try
            {
                /* this.ItemsSyncState already contains items to be synced. */
                var untouched = new HashSet<SyncState<Outlook.AppointmentItem>>(this.ItemsSyncState);
                int nextOffset = -1; // offset of the next page of entries, if any.

                for (int iOffset = 0; iOffset != nextOffset; iOffset = nextOffset)
                {
                    /* get candidates for syncrhonisation from SuiteCRM one page at a time */
                    eGetEntryListResult entriesPage = clsSuiteCRMHelper.GetEntryList(crmModule,
                        String.Format("assigned_user_id = '{0}'", clsSuiteCRMHelper.GetUserId()),
                        0, "date_start DESC", iOffset, false,
                        clsSuiteCRMHelper.GetSugarFields(crmModule));

                    nextOffset = entriesPage.next_offset; // get the offset of the next page

                    if (iOffset != nextOffset)
                    {
                        UpdateItemsFromCrmToOutlook(entriesPage.entry_list, folder, untouched, crmModule);
                    }
                }

                eEntryValue[] invited = clsSuiteCRMHelper.getRelationships("Users",
                    clsSuiteCRMHelper.GetUserId(), crmModule.ToLower(),
                    clsSuiteCRMHelper.GetSugarFields(crmModule));
                if (invited != null)
                {
                    UpdateItemsFromCrmToOutlook(invited, folder, untouched, crmModule);
                }

                try
                {
                    var itemsToBeDeletedFromOutlook = untouched.Where(a => a.ExistedInCrm && a.CrmType == crmModule);
                    foreach (var item in itemsToBeDeletedFromOutlook)
                    {
                        RemoveItemAndSyncState(item);
                    }

                    var itemsToBeAddedToCrm = untouched.Where(a => a.ShouldSyncWithCrm && !a.ExistedInCrm && a.CrmType == crmModule);
                    foreach (var item in itemsToBeAddedToCrm)
                    {
                        AddOrUpdateItemFromOutlookToCrm(item.OutlookItem, crmModule);
                    }
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
        /// Update these appointments 
        /// TODO: This is a candidate for refactoring with ContactSyncing.UpdateItemsFromCrmToOutlook
        /// </summary>
        /// <param name="items">The items to be synchronised.</param>
        /// <param name="folder">The outlook folder to synchronise into.</param>
        /// <param name="untouched">A list of items which have not yet been synchronised; this list is 
        /// modified (destructuvely changed) by the action of this method.</param>
        /// <param name="crmType">The type of CRM objects represented by the appointments.</param>
        private void UpdateItemsFromCrmToOutlook(
            eEntryValue[] items,
            Outlook.MAPIFolder folder, 
            HashSet<SyncState<Outlook.AppointmentItem>> untouched,
            string crmType)
        {
            foreach (var appointment in items)
            {
                try
                {
                    var state = MaybeUpdateAppointmentFromCrmToOutlook(folder, crmType, appointment);
                    if (state != null)
                    {
                        // i.e., the entry was updated...
                        untouched.Remove(state);
                        LogItemAction(state.OutlookItem, "AppointmentSyncing.UpdateAppointmentsFromCrmToOutlook, item removed from untouched");
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("AppointmentSyncing.UpdateAppointmentsFromCrmToOutlook", ex);
                }
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
            dynamic crmItem, 
            DateTime date_start, 
            SyncState<Outlook.AppointmentItem> oItem)
        {
            LogItemAction(oItem.OutlookItem, "AppointmentSyncing.UpdateExistingOutlookItemFromCrm");
            Outlook.AppointmentItem olAppointment = oItem.OutlookItem;
            Outlook.UserProperty olPropertyModifiedDate = olAppointment.UserProperties["SOModifiedDate"];

            if (olPropertyModifiedDate.Value != crmItem.date_modified.value.ToString())
            {
                olAppointment.Subject = crmItem.name.value.ToString();
                olAppointment.Body = crmItem.description.value.ToString();
                if (!string.IsNullOrWhiteSpace(crmItem.date_start.value.ToString()))
                {
                    UpdateOutlookStartAndDuration(crmType, crmItem, date_start, olAppointment);
                }

                SetupSynchronisationPropertiesForOutlookItem(olAppointment, crmType, crmItem);
                olAppointment.Save();
                LogItemAction(oItem.OutlookItem, "AppointmentSyncing.UpdateExistingOutlookItemFromCrm, item saved");
            }
            Log.Warn((string)("Not default dResult.date_modified= " + crmItem.date_modified.value.ToString()));
            oItem.OModifiedDate = DateTime.ParseExact(crmItem.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);

            return oItem;
        }

        /// <summary>
        /// Update this Outlook appointment's start and duration from this CRM object.
        /// </summary>
        /// <param name="crmType">The CRM type of the item from which values are to be taken.</param>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <param name="date_start">The state date/time of the item, adjusted for timezone.</param>
        /// <param name="olAppointment">The outlook item assumed to correspond with the CRM item.</param>
        private void UpdateOutlookStartAndDuration(string crmType, dynamic crmItem, DateTime date_start, Outlook.AppointmentItem olAppointment)
        {
            olAppointment.Start = date_start;
            var minutesString = crmItem.duration_minutes.value.ToString();
            var hoursString = crmItem.duration_hours.value.ToString();

            int minutes = string.IsNullOrWhiteSpace(minutesString) ? 0 : int.Parse(minutesString);
            int hours = string.IsNullOrWhiteSpace(hoursString) ? 0 : int.Parse(hoursString);

            if (crmType == "Meetings")
            {
                olAppointment.Location = crmItem.location.value.ToString();
                olAppointment.End = olAppointment.Start;
                if (hours > 0)
                    olAppointment.End.AddHours(hours);
                if (minutes > 0)
                    olAppointment.End.AddMinutes(minutes);
                Log.Info("\tSetRecepients");
                SetRecipients(olAppointment, crmItem.id.value.ToString(), crmType);
            }
            olAppointment.Duration = minutes + hours * 60;
        }
    }
}
