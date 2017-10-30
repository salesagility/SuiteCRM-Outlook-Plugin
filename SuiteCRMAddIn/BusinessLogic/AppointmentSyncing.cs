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

using System.Diagnostics;

namespace SuiteCRMAddIn.BusinessLogic
{
    using Extensions;
    using ProtoItems;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using Outlook = Microsoft.Office.Interop.Outlook;

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

        /// <summary>
        /// The name of the organiser synchronisation property
        /// </summary>
        public const string OrganiserPropertyName = "SOrganiser";

        /// <summary>
        /// A cache of email addresses to CRM modules and identities
        /// </summary>
        private Dictionary<String, List<AddressResolutionData>> meetingRecipientsCache = 
            new Dictionary<string, List<AddressResolutionData>>();

        public AppointmentSyncing(string name, SyncContext context)
            : base(name, context)
        {
            this.fetchQueryPrefix = "assigned_user_id = '{0}'";
        }

        /// <summary>
        /// Get the id of the record with the specified `smtpAddress` in the module with the specified `moduleName`.
        /// </summary>
        /// <param name="smtpAddress">The SMTP email address to be sought.</param>
        /// <param name="moduleName">The name of the module in which to seek it.</param>
        /// <returns>The corresponding id, if present, else the empty string.</returns>
        public string GetID(string smtpAddress, string moduleName)
        {
            StringBuilder bob = new StringBuilder( $"({moduleName.ToLower()}.id in ")
                .Append( $"(select eabr.bean_id from email_addr_bean_rel eabr ")
                .Append( $"INNER JOIN email_addresses ea on eabr.email_address_id = ea.id ")
                .Append( $"where eabr.bean_module = '{moduleName}' ")
                .Append( $"and ea.email_address LIKE '%{SuiteCRMAddIn.clsGlobals.MySqlEscape(smtpAddress)}%'))");

            string query = bob.ToString();

            Log.Debug($"AppointmentSyncing.GetID: query = `{query}`");

            string[] fields = { "id" };
            EntryList _result = RestAPIWrapper.GetEntryList(moduleName, query, Properties.Settings.Default.SyncMaxRecords, "date_entered DESC", 0, false, fields);

            return _result.result_count > 0 ?
                RestAPIWrapper.GetValueByKey(_result.entry_list[0], "id") :
                string.Empty;
        }

        override public Outlook.MAPIFolder GetDefaultFolder()
        {
            return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
        }

        public override SyncDirection.Direction Direction => Properties.Settings.Default.SyncCalendar;

        protected override bool IsCurrentView => Context.CurrentFolderItemType == Outlook.OlItemType.olAppointmentItem;


        protected override void SaveItem(Outlook.AppointmentItem olItem)
        {
            try
            {
                olItem.Save();
                LogItemAction(olItem, "AppointmentSyncing.SaveItem, saved item");
            }
            catch (System.Exception any)
            {
                Log.Error($"Error while saving appointment {olItem?.Subject}", any);
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
            if (this.permissionsCache.HasExportAccess(AppointmentSyncing.AltCrmModule))
            {
                SyncFolder(GetDefaultFolder(), AppointmentSyncing.AltCrmModule);
            }
            else
            {
                Log.Debug($"AppointmentSyncing.SynchroniseAll: not synchronising {AppointmentSyncing.AltCrmModule} because export access is denied");
            }
        }

        public override string DefaultCrmModule
        {
            get
            {
                return AppointmentSyncing.CrmModule;
            }
        }

        /// <summary>
        /// The actual transmission lock object of this synchroniser.
        /// </summary>
        private object txLock = new object();

        /// <summary>
        /// Allow my parent class to access my transmission lock object.
        /// </summary>
        protected override object TransmissionLock
        {
            get
            {
                return txLock;
            }
        }

        /// <summary>
        /// Ensure that this Outlook item has a property of this name with this value.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        protected override void EnsureSynchronisationPropertyForOutlookItem(Outlook.AppointmentItem olItem, string name, string value)
        {
            try
            {
                Outlook.UserProperty olProperty = olItem.UserProperties[name];
                if (olProperty == null)
                {
                    olProperty = olItem.UserProperties.Add(name, Outlook.OlUserPropertyType.olText);
                }
                if (!olProperty.Value.Equals(value))
                {
                    try
                    {
                        olProperty.Value = value ?? string.Empty;
                        Log.Debug($"AppointmentSyncing.EnsureSynchronisationPropertyForOutlookItem: Set property {name} to value {value} on item {olItem.Subject}");
                    }
                    finally
                    {
                        this.SaveItem(olItem);
                    }
                }
            }
            catch (Exception any)
            {
                Log.Error($"AppointmentSyncing.EnsureSynchronisationPropertyForOutlookItem: Failed to set property {name} to value {value} on item {olItem.Subject}", any);
            }
        }

        protected override void OtherIterationActions()
        {
            CheckMeetingAcceptances();
        }


        /// <summary>
        /// Check meeting acceptances for all future meetings.
        /// </summary>
        private int CheckMeetingAcceptances()
        {
            int result = 0;

            foreach (AppointmentSyncState state in this.ItemsSyncState)
            {
                Outlook.AppointmentItem item = state.OutlookItem;

                if (item.UserProperties[OrganiserPropertyName]?.Value == RestAPIWrapper.GetUserId() &&
                    item.Start > DateTime.Now)
                {
                    result += CheckMeetingAcceptances(item);
                }
            }

            return result;
        }


        /// <summary>
        /// Check meeting acceptances for the invitees of the meeting associated with this `appointment`.
        /// </summary>
        /// <param name="appointment">The appointment.</param>
        /// <returns>the number of valid acceptance statuses found.</returns>
        private int CheckMeetingAcceptances(Outlook.AppointmentItem appointment)
        {
            int count = 0;

            if ( appointment != null) {
                foreach (Outlook.Recipient invitee in appointment.Recipients)
                {
                    string acceptance = string.Empty;

                    switch (invitee.MeetingResponseStatus)
                    {
                        case Outlook.OlResponseStatus.olResponseAccepted:
                            acceptance = "Accept";
                            break;
                        case Outlook.OlResponseStatus.olResponseTentative:
                            acceptance = "Tentative";
                            break;
                        case Microsoft.Office.Interop.Outlook.OlResponseStatus.olResponseNone:
                        case Microsoft.Office.Interop.Outlook.OlResponseStatus.olResponseOrganized:
                        case Microsoft.Office.Interop.Outlook.OlResponseStatus.olResponseNotResponded:
                            // nothing to do
                            break;
                        case Outlook.OlResponseStatus.olResponseDeclined:
                        default:
                            acceptance = "Decline";
                            break;
                    }

                    if (!string.IsNullOrEmpty(acceptance))
                    {
                        this.AddOrUpdateMeetingAcceptanceFromOutlookToCRM(appointment, invitee, acceptance);
                        count++;
                    }
                }
            }

            return count;
        }


        /// <summary>
        /// Check meeting acceptances for the invitees of this `meeting`.
        /// </summary>
        /// <param name="meeting">The meeting.</param>
        /// <returns>the number of valid acceptance statuses found.</returns>
        public int UpdateMeetingAcceptances(Outlook.MeetingItem meeting)
        {
            return meeting == null ? 
                0 : 
                this.CheckMeetingAcceptances(meeting.GetAssociatedAppointment(false));
        }


        /// <summary>
        /// Set the meeting acceptance status, in CRM, for this invitee to this meeting from
        /// their acceptance status in Outlook.
        /// </summary>
        /// <param name="meeting">The appointment item representing the meeting</param>
        /// <param name="invitee">The recipient item representing the invitee</param>
        /// <param name="acceptance">The acceptance status of this invitee of this meeting 
        /// as a string recognised by CRM.</param>
        private void AddOrUpdateMeetingAcceptanceFromOutlookToCRM(Outlook.AppointmentItem meeting, Outlook.Recipient invitee, string acceptance)
        {
            // We don't know which CRM module the invitee belongs to - could be contacts, users, 
            // or indirected via accounts - see AddMeetingRecipientsFromOutlookToCrm. We
            // cannot look this up every time. Therefore we use a cache.
            if (this.meetingRecipientsCache.ContainsKey(invitee.GetSmtpAddress()))
            {
               var meetingId = meeting.UserProperties[CrmIdPropertyName]?.Value;

                if (meetingId != null)
                {
                    foreach (AddressResolutionData resolution in this.meetingRecipientsCache[invitee.GetSmtpAddress()])
                    {
                        RestAPIWrapper.AcceptDeclineMeeting(meetingId.ToString(), resolution.moduleName, resolution.moduleId, acceptance);
                    }
                }
            }
            else
            {
                Log.Warn($"Received {acceptance} to meeting {meeting.Subject} from {invitee.GetSmtpAddress()}, but we have no CRM record for that person");
            }
        }

        private void AddMeetingRecipientsFromOutlookToCrm(Outlook.AppointmentItem olItem, string meetingId)
        {
            LogItemAction(olItem, "AppointmentSyncing.AddMeetingRecipientsFromOutlookToCrm");
            foreach (Outlook.Recipient recipient in olItem.Recipients)
            {
                var smtpAddress = recipient.GetSmtpAddress();

                Log.Info($"recepientName= {recipient.Name}, recepient= {smtpAddress}");

                if (this.meetingRecipientsCache.ContainsKey(smtpAddress))
                {
                    List<AddressResolutionData> resolutions = this.meetingRecipientsCache[smtpAddress];

                    foreach (AddressResolutionData resolution in resolutions)
                    {
                        SetCrmRelationshipFromOutlook(meetingId, resolution);
                    }
                }
                else
                {
                    TryAddRecipientInModule("Leads", meetingId, recipient);
                    TryAddRecipientInModule("Users", meetingId, recipient);
                    TryAddRecipientInModule(ContactSyncing.CrmModule, meetingId, recipient);
                }
            }
        }


        private bool TryAddRecipientInModule(string moduleName, string meetingId, Outlook.Recipient recipient)
        {
            bool result;
            string id = SetCrmRelationshipFromOutlook(meetingId, recipient, moduleName);

            if (!string.IsNullOrWhiteSpace(id))
            {
                string smtpAddress = recipient.GetSmtpAddress();

                this.CacheAddressResolutionData(
                    new AddressResolutionData(moduleName, id, smtpAddress));

                string accountId = RestAPIWrapper.GetRelationship(ContactSyncing.CrmModule, id, "accounts");

                if (!string.IsNullOrWhiteSpace(accountId) &&
                    SetCrmRelationshipFromOutlook(meetingId, "Accounts", accountId))
                {
                    this.CacheAddressResolutionData( 
                        new AddressResolutionData("Accounts", accountId, smtpAddress));
                }

                result = true;
            }
            else
            {
                result = false;
            }

            return result;
        }


        /// <summary>
        /// If a meeting was created in another Outlook we should NOT sync it with CRM because if we do we'll create 
        /// duplicates. Only the Outlook which created it should sync it.
        /// </summary>
        /// <param name="folder">The folder to synchronise into.</param>
        /// <param name="crmType">The CRM type of the candidate item.</param>
        /// <param name="crmItem">The candidate item from CRM.</param>
        /// <returns>True if it's offered to us by CRM with its Outlook ID already populated.</returns>
        protected override bool ShouldAddOrUpdateItemFromCrmToOutlook(Outlook.MAPIFolder folder, string crmType, EntryValue crmItem)
        {
            var outlookId = crmItem.GetValueAsString("outlook_id");
            /* we're good if it's a meeting... */
            bool result = crmType == this.DefaultCrmModule;
            /* provided it doesn't already have an Outlook id */
            result &= string.IsNullOrWhiteSpace(outlookId);
            /* and we're also good if it's an appointment; */
            result |= crmType == AppointmentSyncing.AltCrmModule;
            /* and we're also good if we've already got it */
            result |= (this.GetExistingSyncState(crmItem) != null);

            if (!result)
            {
                Log.Debug($"ShouldAddOrUpdateItemFromCrmToOutlook: not syncing meeting `{crmItem.GetValueAsString("name")}` as it appears to originate from another Outlook instance.");
            }

            return result;
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
            EntryValue crmItem,
            DateTime date_start)
        {
            AppointmentSyncState newState = null;
            Outlook.AppointmentItem olItem = null;
            try
            {
                var crmId = crmItem.GetValueAsString("id");

                /*
                 * There's a nasty little bug (#223) where Outlook offers us back in a different thread
                 * the item we're creating, before we're able to set up the sync state which marks it
                 * as already known. By locking on the enqueueing lock here, we should prevent that.
                 */
                lock (enqueueingLock)
                {
                    olItem = appointmentsFolder.Items.Add(Outlook.OlItemType.olAppointmentItem);

                    olItem.Subject = crmItem.GetValueAsString("name");
                    olItem.Body = crmItem.GetValueAsString("description");
                    /* set the SEntryID property quickly, create the sync state and save the item, to reduce howlaround */
                    EnsureSynchronisationPropertiesForOutlookItem(olItem, crmItem, crmType);

                    this.AddOrGetSyncState(olItem);
                }

                LogItemAction(olItem, "AppointmentSyncing.AddNewItemFromCrmToOutlook");
                if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("date_start")))
                {
                    olItem.Start = date_start;
                    SetOutlookItemDuration(crmType, crmItem, olItem);

                    Log.Info("\tdefault SetRecepients");
                    SetRecipients(olItem, crmId, crmType);
                }
            }
            finally
            {
                if (olItem != null)
                {
                    this.AddOrGetSyncState(olItem);
                }
            }

            return newState;
        }


        /// <summary>
        /// Set this outlook item's duration, but also end time and location, from this CRM item.
        /// </summary>
        /// <param name="crmType">The type of the CRM item.</param>
        /// <param name="crmItem">The CRM item.</param>
        /// <param name="olItem">The Outlook item.</param>
        private void SetOutlookItemDuration(string crmType, EntryValue crmItem, Outlook.AppointmentItem olItem)
        {
            int minutes = 0, hours = 0;
            try
            {
                if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("duration_minutes")))
                {
                    minutes = int.Parse(crmItem.GetValueAsString("duration_minutes"));
                }
                if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("duration_hours")))
                {
                    hours = int.Parse(crmItem.GetValueAsString("duration_hours"));
                }

                int durationMinutes = minutes + hours * 60;

                if (crmType == AppointmentSyncing.CrmModule)
                {
                    olItem.Location = crmItem.GetValueAsString("location");
                    olItem.End = olItem.Start.AddMinutes(durationMinutes);
                }

                olItem.Duration = durationMinutes;
            }
            catch (Exception any)
            {
                Log.Error("AppointmentSyncing.SetOutlookItemDuration", any);
            }
            finally
            {
                this.SaveItem(olItem);
            }
        }

        /// <summary>
        /// Specialisation: in addition to the standard properties, meetings also require an organiser property.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmItem">The CRM item.</param>
        /// <param name="type">The value for the SType property (CRM module name).</param>
        protected override void EnsureSynchronisationPropertiesForOutlookItem(Outlook.AppointmentItem olItem, EntryValue crmItem, string type)
        {
            base.EnsureSynchronisationPropertiesForOutlookItem(olItem, crmItem, type);
            if (this.DefaultCrmModule.Equals(type))
            {
                this.EnsureSynchronisationPropertyForOutlookItem(olItem, OrganiserPropertyName, crmItem.GetValueAsString("assigned_user_id"));
            }
        }

        /// <summary>
        /// Add the item implied by this SyncState, which may not exist in CRM, to CRM.
        /// </summary>
        /// <param name="syncState">The sync state.</param>
        /// <returns>The id of the entry added or updated.</returns>
        internal override string AddOrUpdateItemFromOutlookToCrm(SyncState<Outlook.AppointmentItem> syncState)
        {
            Outlook.AppointmentItem olItem = syncState.OutlookItem;
            Outlook.UserProperty olPropertyType = olItem.UserProperties[TypePropertyName];
            var itemType = olPropertyType != null ? olPropertyType.Value.ToString() : this.DefaultCrmModule;

            return this.AddOrUpdateItemFromOutlookToCrm(syncState, itemType, syncState.CrmEntryId);
        }

        /// <summary>
        /// Add the Outlook item referenced by this sync state, which may not exist in CRM, to CRM.
        /// </summary>
        /// <param name="syncState">The sync state referencing the outlook item to add.</param>
        /// <param name="crmType">The CRM type ('module') to which it should be added</param>
        /// <param name="entryId">The id of this item in CRM, if known (in which case I should be doing
        /// an update, not an add).</param>
        /// <returns>The id of the entry added o</returns>
        internal override string AddOrUpdateItemFromOutlookToCrm(SyncState<Outlook.AppointmentItem> syncState, string crmType, string entryId = "")
        {
            string result = entryId;

            Outlook.AppointmentItem olItem = syncState.OutlookItem;

            if (this.ShouldAddOrUpdateItemFromOutlookToCrm(olItem, crmType))
            {
                if (ShouldDeleteFromCrm(olItem))
                {
                    LogItemAction(olItem, "AppointmentSyncing.AddOrUpdateItemFromOutlookToCrm: Deleting");
 
                    DeleteFromCrm(olItem);
                }
                else if (ShouldDespatchToCrm(olItem))
                {
                    lock (enqueueingLock)
                    {
                        result = base.AddOrUpdateItemFromOutlookToCrm(syncState, crmType, entryId);

                        if (String.IsNullOrEmpty(result))
                        {
                            Log.Warn("AppointmentSyncing.AddOrUpdateItemFromOutlookToCrm: Invalid CRM Id returned; item may not have been stored.");
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(entryId))
                            {
                                /* i.e. this was a new item saved to CRM for the first time */
                                SetCrmRelationshipFromOutlook(result, "Users", RestAPIWrapper.GetUserId());

                                this.SaveItem(olItem);

                                if (olItem.Recipients != null)
                                {
                                    AddMeetingRecipientsFromOutlookToCrm(olItem, result);
                                }
                            }
                        }
                    }
                }
                else
                {
                    LogItemAction(olItem, "AppointmentSyncing.AddItemFromOutlookToCrm, Not despatching");
                }
            }
            else
            {
                LogItemAction(olItem, "AppointmentSyncing.AddItemFromOutlookToCrm, Not enabled");
            }

            return result;
        }


        /// <summary>
        /// Construct a JSON packet representing this Outlook item, and despatch it to CRM. 
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmType">The type within CRM to which the item should be added.</param>
        /// <param name="entryId">The corresponding entry id in CRM, if known.</param>
        /// <returns>The CRM id of the object created or modified.</returns>
        protected override string ConstructAndDespatchCrmItem(Outlook.AppointmentItem olItem, string crmType, string entryId)
        {
            return RestAPIWrapper.SetEntryUnsafe(new ProtoAppointment(olItem).AsNameValues(entryId), crmType);
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
                SyncState<Outlook.AppointmentItem> syncStateForItem = GetExistingSyncState(olItem);
                if (syncStateForItem != null)
                {
                    this.RemoveFromCrm(syncStateForItem);
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
                foreach (Outlook.AppointmentItem olItem in appointmentsFolder.Items)
                {
                    if (olItem.Start >= this.GetStartDate())
                    {
                        Outlook.UserProperty olPropertyModified = olItem.UserProperties[ModifiedDatePropertyName];
                        Outlook.UserProperty olPropertyType = olItem.UserProperties[TypePropertyName];
                        Outlook.UserProperty olPropertyEntryId = olItem.UserProperties[CrmIdPropertyName];
                        if (olPropertyModified != null &&
                            olPropertyType != null &&
                            olPropertyEntryId != null)
                        {
                            /* The appointment probably already has the three magic properties 
                             * required for synchronisation; is that a proxy for believing that it
                             * already exists in CRM? If so, is it reliable? */
                            LogItemAction(olItem, "AppointmentSyncing.GetOutlookItems: Adding known item to queue");
                        }
                        else
                        {
                            LogItemAction(olItem, "AppointmentSyncing.GetOutlookItems: Adding unknown item to queue");
                        }

                        this.AddOrGetSyncState(olItem);
                    }
                }                
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.GetOutlookCalItems", ex);
            }
        }

        /// <summary>
        /// Log a message regarding this Outlook appointment.
        /// </summary>
        /// <param name="olItem">The outlook item.</param>
        /// <param name="message">The message to be logged.</param>
        internal override void LogItemAction(Outlook.AppointmentItem olItem, string message)
        {
            try
            {
                Outlook.UserProperty olPropertyEntryId = olItem.UserProperties[CrmIdPropertyName];
                string crmId = olPropertyEntryId == null ?
                    "[not present]" :
                    olPropertyEntryId.Value;
                StringBuilder bob = new StringBuilder();
                bob.Append($"{message}:\n\tOutlook Id  : {olItem.EntryID}\n\tCRM Id      : {crmId}\n\tSubject     : '{olItem.Subject}'\n\tSensitivity : {olItem.Sensitivity}\n\tRecipients:\n");
                foreach (Outlook.Recipient recipient in olItem.Recipients)
                {
                    bob.Append($"\t\t{recipient.Name}: {recipient.GetSmtpAddress()}\n");
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
        /// <param name="crmItem">The candidate item from CRM.</param>
        /// <returns>The synchronisation state of the item updated (if it was updated).</returns>
        protected override SyncState<Outlook.AppointmentItem> AddOrUpdateItemFromCrmToOutlook(
            Outlook.MAPIFolder folder,
            string crmType,
            EntryValue crmItem)
        {
            SyncState<Outlook.AppointmentItem> result = null;
            DateTime dateStart = crmItem.GetValueAsDateTime("date_start");

            if (dateStart >= GetStartDate())
            {
                /* search for the item among the sync states I already know about */
                var syncState = this.GetExistingSyncState(crmItem);
                if (syncState == null)
                {
                    /* check for howlaround */
                    var matches = this.FindMatches(crmItem);

                    if (matches.Count == 0)
                    {
                        /* didn't find it, so add it to Outlook */
                        result = AddNewItemFromCrmToOutlook(folder, crmType, crmItem, dateStart);
                    }
                    else
                    {
                        this.Log.Warn($"Howlaround detected? Appointment '{crmItem.GetValueAsString("name")}' offered with id {crmItem.GetValueAsString("id")}, expected {matches[0].CrmEntryId}, {matches.Count} duplicates");
                    }
                }
                else
                {
                    /* found it, so update it from the CRM item */
                    result = UpdateExistingOutlookItemFromCrm(crmType, crmItem, dateStart, syncState);

                    result?.OutlookItem.Save();
                }

                if (crmItem?.relationships?.link_list != null)
                {
                    foreach (var list in crmItem.relationships.link_list)
                    {
                        foreach (var record in list.records)
                        {
                            var data = record.data.AsDictionary();
                            try
                            {
                                this.CacheAddressResolutionData(list.name, record);
                            }
                            catch (KeyNotFoundException kex)
                            {
                                Log.Error($"Key not found while caching meeting recipients.", kex);
                            }
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Add an address resolution composed from this module name and record to the cache.
        /// </summary>
        /// <param name="moduleName">The name of the module in which the record was found</param>
        /// <param name="record">The record.</param>
        private void CacheAddressResolutionData(string moduleName, LinkRecord record)
        {
            Dictionary<string, object> data = record.data.AsDictionary();
            string smtpAddress = data[AddressResolutionData.EmailAddressFieldName].ToString();
            AddressResolutionData resolution = new AddressResolutionData(moduleName, data);

            CacheAddressResolutionData(resolution);
        }

        /// <summary>
        /// Add this resolution to the cache.
        /// </summary>
        /// <param name="resolution">The resolution to add.</param>
        private void CacheAddressResolutionData(AddressResolutionData resolution)
        {
            List<AddressResolutionData> resolutions;

            if (this.meetingRecipientsCache.ContainsKey(resolution.emailAddress)) {
                resolutions = this.meetingRecipientsCache[resolution.emailAddress];
            }
            else
            {
                resolutions = new List<AddressResolutionData>();
                this.meetingRecipientsCache[resolution.emailAddress] = resolutions;
            }

            if (!resolutions.Any( x => x.moduleId == resolution.moduleId && x.moduleName == resolution.moduleName))
            {
                resolutions.Add(resolution);
            }

            Log.Debug($"Successfully cached recipient {resolution.emailAddress} => {resolution.moduleName}, {resolution.moduleId}.");
        }

        protected override bool IsMatch(Outlook.AppointmentItem olItem, EntryValue crmItem)
        {
            var crmItemStart = crmItem.GetValueAsDateTime("date_start");
            var crmItemName = crmItem.GetValueAsString("name");

            var olItemStart = olItem.Start;
            var subject = olItem.Subject;

            return subject == crmItemName &&
                olItemStart == crmItemStart;
        }

        /// <summary>
        /// Remove the synchronisation properties from this Outlook item.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        private static void RemoveSynchronisationPropertiesFromOutlookItem(Outlook.AppointmentItem olItem)
        {
            RemoveSynchronisationPropertyFromOutlookItem(olItem, CrmIdPropertyName);
            RemoveSynchronisationPropertyFromOutlookItem(olItem, TypePropertyName);
            RemoveSynchronisationPropertyFromOutlookItem(olItem, ModifiedDatePropertyName);
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
                try
                {
                    olItem.UserProperties.Remove(found);
                }
                finally
                {
                    olItem.Save();
                }
            }
        }


        /// <summary>
        /// Sets up a CRM relationship to mimic an Outlook relationship
        /// </summary>
        /// <param name="meetingId">The ID of the meeting</param>
        /// <param name="recipient">The outlook recipient representing the person to link with.</param>
        /// <param name="foreignModule">the name of the module we're seeking to link with.</param>
        /// <returns>True if a relationship was created.</returns>
        private string SetCrmRelationshipFromOutlook(string meetingId, Outlook.Recipient recipient, string foreignModule)
        {
            string foreignId = GetID(recipient.GetSmtpAddress(), foreignModule);

            return !string.IsNullOrWhiteSpace(foreignId) && 
                SetCrmRelationshipFromOutlook(meetingId, foreignModule, foreignId) ?
                foreignId :
                string.Empty;
        }


        /// <summary>
        /// Sets up a CRM relationship to mimic an Outlook relationship
        /// </summary>
        /// <param name="meetingId">The meeting id.</param>
        /// <param name="resolution">Address resolution data from the cache.</param>
        /// <returns>True if a relationship was created.</returns>
        private bool SetCrmRelationshipFromOutlook(string meetingId, AddressResolutionData resolution)
        {
            return this.SetCrmRelationshipFromOutlook(meetingId, resolution.moduleName, resolution.moduleId);
        }


        /// <summary>
        /// Sets up a CRM relationship to mimic an Outlook relationship
        /// </summary>
        /// <param name="meetingId">The ID of the meeting</param>
        /// <param name="foreignModule">the name of the module we're seeking to link with.</param>
        /// <param name="foreignId">The id in the foreign module of the record we're linking to.</param>
        /// <returns>True if a relationship was created.</returns>
        private bool SetCrmRelationshipFromOutlook(string meetingId, string foreignModule, string foreignId)
        {
            bool result = false;

            if (foreignId != String.Empty)
            {
                SetRelationshipParams info = new SetRelationshipParams
                {
                    module2 = AppointmentSyncing.CrmModule,
                    module2_id = meetingId,
                    module1 = foreignModule,
                    module1_id = foreignId
                };
                result = RestAPIWrapper.SetRelationshipUnsafe(info);
            }

            return result;
        }


        private void SetRecipients(Outlook.AppointmentItem olItem, string sMeetingID, string sModule)
        {
            this.LogItemAction(olItem, "SetRecipients");

            try
            {
                olItem.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                int iCount = olItem.Recipients.Count;
                for (int iItr = 1; iItr <= iCount; iItr++)
                {
                    olItem.Recipients.Remove(1);
                }

                string[] invitee_categories = { "users", ContactSyncing.CrmModule, "leads" };
                foreach (string invitee_category in invitee_categories)
                {
                    EntryValue[] relationships = RestAPIWrapper.GetRelationships(sModule, sMeetingID, invitee_category, new string[] { "id", "email1", "phone_work" });
                    if (relationships != null)
                    {

                        foreach (var relationship in relationships)
                        {
                            string phone_work = relationship.GetValueAsString("phone_work");
                            string email1 = relationship.GetValueAsString("email1");
                            string identifier = (sModule == AppointmentSyncing.CrmModule) || string.IsNullOrWhiteSpace(phone_work) ?
                                    email1 :
                                    $"{email1} : {phone_work}";

                            if (!String.IsNullOrWhiteSpace(identifier))
                            {
                                olItem.Recipients.Add(identifier);
                            }
                        }
                    }
                }
            }
            finally
            {
                this.SaveItem(olItem);
            }
        }

        /// <summary>
        /// We should delete an item from CRM if it already exists in CRM, but it is now private.
        /// </summary>
        /// <param name="olItem">The Outlook item</param>
        /// <returns>true if the Outlook item should be deleted from CRM.</returns>
        private bool ShouldDeleteFromCrm(Outlook.AppointmentItem olItem)
        {
            Outlook.UserProperty olPropertyEntryId = olItem.UserProperties[CrmIdPropertyName];
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
            var syncConfigured = SyncDirection.AllowOutbound(Properties.Settings.Default.SyncCalendar);
            string organiser = olItem.Organizer;
            var currentUser = Application.Session.CurrentUser;
            var exchangeUser = currentUser.AddressEntry.GetExchangeUser();
            var currentUserName = exchangeUser == null ? 
                Application.Session.CurrentUser.Name:
                exchangeUser.Name;
            string crmId = olItem.UserProperties[CrmIdPropertyName]?.Value;

            return olItem != null &&
                syncConfigured && 
                olItem.Sensitivity == Outlook.OlSensitivity.olNormal &&
                /* If there is a valid crmId it's arrived via CRM and is therefore safe to save to CRM;
                 * if the current user is the organiser, AND there's no valid CRM id, then it's a new one
                 * that the current user made, and we should save it to CRM. */
                (!string.IsNullOrEmpty(crmId) || currentUserName == organiser);
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

                EntryValue[] invited = RestAPIWrapper.GetRelationships("Users",
                    RestAPIWrapper.GetUserId(), crmModule.ToLower(),
                    RestAPIWrapper.GetSugarFields(crmModule));
                if (invited != null)
                {
                    AddOrUpdateItemsFromCrmToOutlook(invited, folder, untouched, crmModule);
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
        /// <param name="syncState">The outlook item assumed to correspond with the CRM item.</param>
        /// <returns>An appropriate sync state.</returns>
        private SyncState<Outlook.AppointmentItem> UpdateExistingOutlookItemFromCrm(
            string crmType, 
            EntryValue crmItem, 
            DateTime date_start, 
            SyncState<Outlook.AppointmentItem> syncState)
        {
            LogItemAction(syncState.OutlookItem, "AppointmentSyncing.UpdateExistingOutlookItemFromCrm");

            if (!syncState.IsDeletedInOutlook)
            {
                Outlook.AppointmentItem olItem = syncState.OutlookItem;
                Outlook.UserProperty olPropertyModifiedDate = olItem.UserProperties[ModifiedDatePropertyName];

                if (olPropertyModifiedDate.Value != crmItem.GetValueAsString("date_modified"))
                {
                    try
                    {
                        olItem.Subject = crmItem.GetValueAsString("name");
                        olItem.Body = crmItem.GetValueAsString("description");
                        if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("date_start")))
                        {
                            UpdateOutlookStartAndDuration(crmType, crmItem, date_start, olItem);
                        }

                        EnsureSynchronisationPropertiesForOutlookItem(olItem, crmItem, crmType);
                        LogItemAction(syncState.OutlookItem, "AppointmentSyncing.UpdateExistingOutlookItemFromCrm, item saved");
                    }
                    finally
                    {
                        this.SaveItem(olItem);
                    }
                }
                Log.Warn((string)("Not default dResult.date_modified= " + crmItem.GetValueAsString("date_modified")));
                syncState.OModifiedDate = DateTime.ParseExact(crmItem.GetValueAsString("date_modified"), "yyyy-MM-dd HH:mm:ss", null);
            }

            return syncState;
        }

        /// <summary>
        /// Update this Outlook appointment's start and duration from this CRM object.
        /// </summary>
        /// <param name="crmType">The CRM type of the item from which values are to be taken.</param>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <param name="date_start">The state date/time of the item, adjusted for timezone.</param>
        /// <param name="olItem">The outlook item assumed to correspond with the CRM item.</param>
        private void UpdateOutlookStartAndDuration(string crmType, EntryValue crmItem, DateTime date_start, Outlook.AppointmentItem olItem)
        {
            try
            {
                olItem.Start = date_start;
                var minutesString = crmItem.GetValueAsString("duration_minutes");
                var hoursString = crmItem.GetValueAsString("duration_hours");

                int minutes = string.IsNullOrWhiteSpace(minutesString) ? 0 : int.Parse(minutesString);
                int hours = string.IsNullOrWhiteSpace(hoursString) ? 0 : int.Parse(hoursString);

                if (crmType == AppointmentSyncing.CrmModule)
                {
                    olItem.Location = crmItem.GetValueAsString("location");
                    olItem.End = olItem.Start;
                    if (hours > 0)
                    {
                        olItem.End.AddHours(hours);
                    }
                    if (minutes > 0)
                    {
                        olItem.End.AddMinutes(minutes);
                    }
                    SetRecipients(olItem, crmItem.GetValueAsString("id"), crmType);
                }
                olItem.Duration = minutes + hours * 60;
            }
            finally
            {
                this.SaveItem(olItem);
            }
        }

        protected override SyncState<Outlook.AppointmentItem> ConstructSyncState(Outlook.AppointmentItem oItem)
        {
            return new AppointmentSyncState(oItem.UserProperties[TypePropertyName]?.Value.ToString())
            {
                OutlookItem = oItem,
                CrmEntryId = oItem.UserProperties[CrmIdPropertyName]?.Value.ToString(),
                OModifiedDate = ParseDateTimeFromUserProperty(oItem.UserProperties[ModifiedDatePropertyName]?.Value.ToString()),
            };
        }

        internal override string GetOutlookEntryId(Outlook.AppointmentItem olItem)
        {
            return olItem.EntryID;
        }

        protected override string GetCrmEntryId(Outlook.AppointmentItem olItem)
        {
            return olItem?.UserProperties[CrmIdPropertyName]?.Value.ToString();
        }

        /// <summary>
        /// Return the sensitivity of this outlook item.
        /// </summary>
        /// <remarks>
        /// Outlook item classes do not inherit from a common base class, so generic client code cannot refer to 'OutlookItem.Sensitivity'.
        /// </remarks>
        /// <param name="item">The outlook item whose sensitivity is required.</param>
        /// <returns>the sensitivity of the item.</returns>
        internal override Outlook.OlSensitivity GetSensitivity(Outlook.AppointmentItem item)
        {
            return item.Sensitivity;
        }


        /// <summary>
        /// Used for caching data for resolving email addresses to CRM records.
        /// </summary>
        private class AddressResolutionData
        {
            /// <summary>
            /// Expected name in the input map of the email address field.
            /// </summary>
            public const string EmailAddressFieldName = "email1";

            /// <summary>
            /// Expected name in the input map of the field containing the id in 
            /// the specified module.
            /// </summary>
            public const string ModuleIdFieldName = "id";

            /// <summary>
            /// Expected name in the input map of the field containing an associated id in 
            /// the `Accounts` module, if any.
            /// </summary>
            public const string AccountIdFieldName = "account_id";

            /// <summary>
            /// The email address resolved by this data.
            /// </summary>
            public readonly string emailAddress;
            /// <summary>
            /// The name of the CRM module to which it resolves.
            /// </summary>
            public readonly string moduleName;
            /// <summary>
            /// The id within that module of the record to which it resolves.
            /// </summary>
            public readonly string moduleId;

            /// <summary>
            /// The id within the `Accounts` module of a related record, if any.
            /// </summary>
            private readonly object accountId;

            public AddressResolutionData(string moduleName, string moduleId, string emailAddress)
            {
                this.moduleName = moduleName;
                this.moduleId = moduleId;
                this.emailAddress = emailAddress;
            }

            public AddressResolutionData( string moduleName, Dictionary<string, object> data)
            {
                this.moduleName = moduleName;
                this.moduleId = data[ModuleIdFieldName]?.ToString();
                this.emailAddress = data[EmailAddressFieldName]?.ToString();
                try
                {
                    this.accountId = data[AccountIdFieldName]?.ToString();
                }
                catch (KeyNotFoundException)
                {
                    // and ignore it; that key often won't be there.
                }
            }
        }
    }
}
