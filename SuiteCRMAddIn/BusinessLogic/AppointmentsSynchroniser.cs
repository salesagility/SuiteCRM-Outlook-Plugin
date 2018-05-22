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
    using Exceptions;
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
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Handles the synchronisation of appointments between Outlook and CMS.
    /// </summary>
    public abstract class AppointmentsSynchroniser<SyncStateType> : Synchroniser<Outlook.AppointmentItem, SyncStateType>
        where SyncStateType : SyncState<Outlook.AppointmentItem>
    {
        /// <summary>
        /// The name of the organiser synchronisation property
        /// </summary>
        public const string OrganiserPropertyName = "SOrganiser";

        /// <summary>
        /// Microsoft Conferencing Add-in creates temporary items whose names begin 
        /// 'PLEASE IGNORE'. We should not sync these.
        /// </summary>
        public const string MSConfTmpSubjectPrefix = "PLEASE IGNORE";

        /// <summary>
        /// A cache of email addresses to CRM modules and identities
        /// </summary>
        protected Dictionary<String, List<AddressResolutionData>> meetingRecipientsCache =
            new Dictionary<string, List<AddressResolutionData>>();


        public AppointmentsSynchroniser(string name, SyncContext context)
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
        public string GetInviteeIdBySmtpAddress(string smtpAddress, string moduleName)
        {
            StringBuilder bob = new StringBuilder( $"({moduleName.ToLower()}.id in ")
                .Append( $"(select eabr.bean_id from email_addr_bean_rel eabr ")
                .Append( $"INNER JOIN email_addresses ea on eabr.email_address_id = ea.id ")
                .Append( $"where eabr.bean_module = '{moduleName}' ")
                .Append( $"and ea.email_address LIKE '%{RestAPIWrapper.MySqlEscape(smtpAddress)}%'))");

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

        /// <summary>
        /// #2246: Discriminate between calls and meetings when adding and updating.
        /// </summary>
        protected override void OutlookItemAdded(Outlook.AppointmentItem olItem)
        {
            if (Globals.ThisAddIn.IsLicensed)
            {
                try
                {
                    if (olItem != null && olItem.IsCall())
                    {
                        base.OutlookItemAdded<CallSyncState>(olItem, Globals.ThisAddIn.CallsSynchroniser);
                    }
                    else
                    {
                        base.OutlookItemAdded<MeetingSyncState>(olItem, Globals.ThisAddIn.MeetingsSynchroniser);
                    }
                }
                finally
                {
                    if (olItem != null)
                    {
                        SaveItem(olItem);
                    }
                }
            }
            else
            {
                Log.Warn($"Synchroniser.OutlookItemAdded: item {this.GetOutlookEntryId(olItem)} not added because not licensed");
            }
        }

        /// <summary>
        /// #2246: Discriminate between calls and meetings when adding and updating.
        /// </summary>
        protected override void OutlookItemChanged(Outlook.AppointmentItem olItem)
        {
            if (Globals.ThisAddIn.IsLicensed)
            {
                try
                {
                    if (olItem != null && olItem.IsCall())
                    {
                        base.OutlookItemChanged<CallSyncState>(olItem, Globals.ThisAddIn.CallsSynchroniser);
                    }
                    else
                    {
                        base.OutlookItemChanged<MeetingSyncState>(olItem, Globals.ThisAddIn.MeetingsSynchroniser);
                    }
                }
                catch (BadStateTransition bst)
                {
                    Log.Warn("Bad state transition in OutlookItemChanged - if transition Transmitted => Pending fails that's OK", bst);
                    /* couldn't set pending -> transmission is in progress */
                }
                finally
                {
                    this.SaveItem(olItem);
                }
            }
            else
            {
                Log.Warn($"Synchroniser.OutlookItemAdded: item {this.GetOutlookEntryId(olItem)} not updated because not licensed");
            }
        }

        protected override void SaveItem(Outlook.AppointmentItem olItem)
        {
            try
            {
                olItem.Save();
                LogItemAction(olItem, "AppointmentSyncing.SaveItem, saved item");
            }
            catch (System.Exception any)
            {
                try
                {
                    Log.Error($"Error while saving appointment {olItem?.Subject}", any);
                }
                catch (COMException comx)
                {
                    Log.Error($"COM exception while trying to save appointment, appointment has probably been deleted.", comx);
                }
            }
        }

        /// <summary>
        /// Prefix for meetings which have been canceled.
        /// </summary>
        private static readonly string CanceledPrefix = "CANCELED";


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
                    catch (Exception any)
                    {
                        Log.Error($"AppointmentSyncing.EnsureSynchronisationPropertyForOutlookItem: Failed to set property {name} to value {value} on item {olItem.Subject}", any);
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
            /* and we're also good if we've already got it */
            result |= (SyncStateManager.Instance.GetExistingSyncState(crmItem) != null);

            if (!result)
            {
                Log.Debug($"ShouldAddOrUpdateItemFromCrmToOutlook: not syncing meeting `{crmItem.GetValueAsString("name")}` as it appears to originate from another Outlook instance.");
            }

            return result;
        }


        /// <summary>
        /// Add an item existing in CRM but not found in Outlook to Outlook.
        /// </summary>
        /// <remarks>
        /// This method is disconcertingly different from equivalent methods in other synchronisers;
        /// TODO: the differences ought to be thought about.
        /// </remarks>
        /// <see cref="ContactSynchroniser.AddNewItemFromCrmToOutlook(Outlook.MAPIFolder, EntryValue)"/> 
        /// <param name="appointmentsFolder">The Outlook folder in which the item should be stored.</param>
        /// <param name="crmType">The CRM type of the item from which values are to be taken.</param>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <param name="date_start">The state date/time of the item, adjusted for timezone.</param>
        /// <returns>A sync state object for the new item.</returns>
        protected virtual SyncStateType AddNewItemFromCrmToOutlook(
            Outlook.MAPIFolder appointmentsFolder,
            string crmType,
            EntryValue crmItem,
            DateTime date_start)
        {
            SyncStateType newState = null;
            Outlook.AppointmentItem olItem = null;

            Log.Debug(
                (string)string.Format(
                    $"{this.GetType().Name}.AddNewItemFromCrmToOutlook, entry id is '{crmItem.GetValueAsString("id")}', creating in Outlook."));

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
                    SetMeetingStatus(olItem, crmItem);

                    /* set the SEntryID property quickly, create the sync state and save the item, to reduce howlaround */
                    EnsureSynchronisationPropertiesForOutlookItem(olItem, crmItem, crmType);
                }

                LogItemAction(olItem, "AppointmentSyncing.AddNewItemFromCrmToOutlook");
                if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("date_start")))
                {
                    olItem.Start = date_start;
                    SetOutlookItemDuration(crmType, crmItem, olItem);
                }
            }
            catch (Exception any)
            {
                Log.Warn("Unexpected error in AppointmentSyncing.AddNewItemFromCrmToOutlook", any);
            }
            finally
            {
                if (olItem != null)
                {
                    newState = SyncStateManager.Instance.GetOrCreateSyncState(olItem) as SyncStateType;
                    newState.SetNewFromCRM();
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
        protected void SetOutlookItemDuration(string crmType, EntryValue crmItem, Outlook.AppointmentItem olItem)
        {
            try
            {
                SetOutlookItemDuration(crmItem, olItem);
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
        /// Set this outlook item's duration from this CRM item.
        /// </summary>
        /// <param name="crmType">The type of the CRM item.</param>
        /// <param name="crmItem">The CRM item.</param>
        /// <param name="olItem">The Outlook item.</param>
        protected virtual void SetOutlookItemDuration(EntryValue crmItem, Outlook.AppointmentItem olItem)
        {
            int minutes = 0, hours = 0;

            if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("duration_minutes")))
            {
                minutes = int.Parse(crmItem.GetValueAsString("duration_minutes"));
            }
            if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("duration_hours")))
            {
                hours = int.Parse(crmItem.GetValueAsString("duration_hours"));
            }

            int durationMinutes = minutes + hours * 60;

            olItem.Duration = durationMinutes;
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
        internal override string AddOrUpdateItemFromOutlookToCrm(SyncState<Outlook.AppointmentItem> syncState, string entryId = "")
        {
            Outlook.AppointmentItem olItem = syncState.OutlookItem;

            try
            {
                Outlook.UserProperty olPropertyType = olItem.UserProperties[SyncStateManager.TypePropertyName];
                Outlook.UserProperty olPropertyCrmId = olItem.UserProperties[SyncStateManager.CrmIdPropertyName];

                string result = string.Empty;

                if (this.ShouldAddOrUpdateItemFromOutlookToCrm(olItem))
                {
                    if (ShouldDeleteFromCrm(olItem))
                    {
                        LogItemAction(olItem, "AppointmentSyncing.AddOrUpdateItemFromOutlookToCrm: Deleting");

                        DeleteFromCrm(olItem);
                    }
                    else if (ShouldDespatchToCrm(olItem))
                    {
                        result = base.AddOrUpdateItemFromOutlookToCrm(syncState);

                        if (String.IsNullOrEmpty(result))
                        {
                            Log.Warn("AppointmentSyncing.AddOrUpdateItemFromOutlookToCrm: Invalid CRM Id returned; item may not have been stored.");
                        }
                        else if (olPropertyCrmId == null || string.IsNullOrEmpty(olPropertyCrmId.Value))
                        {
                            /* i.e. this was a new item saved to CRM for the first time */
                            if (syncState.OutlookItem.IsCall())
                            {
                                SetCrmRelationshipFromOutlook(Globals.ThisAddIn.CallsSynchroniser, result, "Users", RestAPIWrapper.GetUserId());
                            }
                            else
                            {
                                SetCrmRelationshipFromOutlook(Globals.ThisAddIn.MeetingsSynchroniser, result, "Users", RestAPIWrapper.GetUserId());
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
            catch (COMException)
            {
                this.HandleItemMissingFromOutlook(syncState);
                return syncState.CrmEntryId;
            }
        }


        /// <summary>
        /// Construct a JSON packet representing this Outlook item, and despatch it to CRM. 
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="entryId">The corresponding entry id in CRM, if known.</param>
        /// <returns>The CRM id of the object created or modified.</returns>
        protected override string ConstructAndDespatchCrmItem(Outlook.AppointmentItem olItem, string entryId)
        {
            return RestAPIWrapper.SetEntryUnsafe(new ProtoAppointment<SyncStateType>(olItem).AsNameValues(entryId), this.DefaultCrmModule);
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
                SyncState<Outlook.AppointmentItem> syncStateForItem = SyncStateManager.Instance.GetExistingSyncState(olItem);
                if (syncStateForItem != null)
                {
                    this.RemoveFromCrm(syncStateForItem);
                    this.RemoveItemSyncState(syncStateForItem);
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
                List<Outlook.AppointmentItem> deletionCandidates = new List<Outlook.AppointmentItem>();
                foreach (Outlook.AppointmentItem olItem in appointmentsFolder.Items)
                {
                    try
                    {
                        if (olItem.Start >= this.GetStartDate())
                        {
                            Outlook.UserProperty olPropertyModified = olItem.UserProperties[SyncStateManager.ModifiedDatePropertyName];
                            Outlook.UserProperty olPropertyType = olItem.UserProperties[SyncStateManager.TypePropertyName];
                            Outlook.UserProperty olPropertyEntryId = olItem.UserProperties[SyncStateManager.CrmIdPropertyName];
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

                            SyncStateManager.Instance.GetOrCreateSyncState(olItem).SetPresentAtStartup();
                        }
                    }
                    catch (ProbableDuplicateItemException<Outlook.AppointmentItem>)
                    {
                        deletionCandidates.Add(olItem);
                    }
                }
                
                foreach (var toDelete in deletionCandidates)
                {
                    toDelete.Delete();
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
                Outlook.UserProperty olPropertyEntryId = olItem.UserProperties[SyncStateManager.CrmIdPropertyName];
                string crmId = olPropertyEntryId == null ?
                    "[not present]" :
                    olPropertyEntryId.Value;
                StringBuilder bob = new StringBuilder();
                bob.Append($"{message}:\n\tOutlook Id  : {olItem.EntryID}")
                    .Append($"\n\tGlobal Id   : {olItem.GlobalAppointmentID}")
                    .Append($"\n\tCRM Id      : {crmId}")
                    .Append($"\n\tSubject     : '{olItem.Subject}'")
                    .Append($"\n\tSensitivity : {olItem.Sensitivity}")
                    .Append($"\n\tStatus      : {olItem.MeetingStatus}")
                    .Append($"\n\tReminder set: {olItem.ReminderSet}")
                    .Append($"\n\tOrganiser   : {olItem.Organizer}")
                    .Append($"\n\tOutlook User: {clsGlobals.GetCurrentUsername()}")
                    .Append($"\n\tTxState     : {SyncStateManager.Instance.GetExistingSyncState(olItem)?.TxState}")
                    .Append($"\n\tRecipients  :\n");

                foreach (Outlook.Recipient recipient in olItem.Recipients)
                {
                    bob.Append($"\t\t{recipient.Name}: {recipient.GetSmtpAddress()} - ({recipient.MeetingResponseStatus})\n");
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
            SyncState existing = SyncStateManager.Instance.GetExistingSyncState(crmItem);
            SyncStateType result = existing as SyncStateType;

            DateTime dateStart = crmItem.GetValueAsDateTime("date_start");

            if (dateStart >= GetStartDate())
            {
                /* search for the item among the sync states I already know about */
                if (existing == null)
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
                else if (result != null)
                {
                    /* found it, so update it from the CRM item */
                    UpdateExistingOutlookItemFromCrm(crmType, crmItem, dateStart, result);
                }
                else
                {
                    throw new UnexpectedSyncStateClassException($"{this.GetType().Name}", existing);
                }

                existing?.SaveItem();
            }

            return result;
        }

        internal override void HandleItemMissingFromOutlook(SyncState<Outlook.AppointmentItem> syncState)
        {
            if (syncState.CrmType == MeetingsSynchroniser.CrmModule)
            {
                /* typically, when this method is called, the Outlook Item will already be invalid, and if it is not,
                 * it may become invalid during the execution of this method. So this method CANNOT depend on any
                 * values taken from the Outlook item. */
                EntryList entries = RestAPIWrapper.GetEntryList(
                    syncState.CrmType, $"id = {syncState.CrmEntryId}",
                    Properties.Settings.Default.SyncMaxRecords,
                    "date_entered DESC", 0, false, null);

                if (entries.entry_list.Count() > 0)
                {
                    this.HandleItemMissingFromOutlook(entries.entry_list[0], syncState, syncState.CrmType);
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
        protected string SetCrmRelationshipFromOutlook<T, S>(Synchroniser<T, S> sync, string meetingId, Outlook.Recipient recipient, string foreignModule)
            where T : class
            where S : SyncState<T>
        {
            string foreignId = GetInviteeIdBySmtpAddress(recipient.GetSmtpAddress(), foreignModule);

            return !string.IsNullOrWhiteSpace(foreignId) &&
                SetCrmRelationshipFromOutlook(sync, meetingId, foreignModule, foreignId) ?
                foreignId :
                string.Empty;
        }


        /// <summary>
        /// Sets up a CRM relationship to mimic an Outlook relationship
        /// </summary>
        /// <param name="meetingId">The meeting id.</param>
        /// <param name="resolution">Address resolution data from the cache.</param>
        /// <returns>True if a relationship was created.</returns>
        protected bool SetCrmRelationshipFromOutlook<T, S>(Synchroniser<T, S> sync, string meetingId, AddressResolutionData resolution)
            where T : class
            where S : SyncState<T>
        {
            return this.SetCrmRelationshipFromOutlook<T, S>(sync, meetingId, resolution.moduleName, resolution.moduleId);
        }


        /// <summary>
        /// Sets up a CRM relationship to mimic an Outlook relationship
        /// </summary>
        /// <param name="meetingId">The ID of the meeting</param>
        /// <param name="foreignModule">the name of the module we're seeking to link with.</param>
        /// <param name="foreignId">The id in the foreign module of the record we're linking to.</param>
        /// <returns>True if a relationship was created.</returns>
        protected bool SetCrmRelationshipFromOutlook<T, S>(Synchroniser<T, S> sync, string meetingId, string foreignModule, string foreignId)
            where T : class
            where S : SyncState<T>
        {
            bool result = false;

            if (foreignId != String.Empty)
            {
                SetRelationshipParams info = new SetRelationshipParams
                {
                    module2 = sync.DefaultCrmModule,
                    module2_id = meetingId,
                    module1 = foreignModule,
                    module1_id = foreignId
                };
                result = RestAPIWrapper.SetRelationshipUnsafe(info);
            }

            return result;
        }


        /// <summary>
        /// Override: we get notified of a removal, for a Meeting item, when the meeting is
        /// cancelled. We do NOT want to remove such an item; instead, we want to update it.
        /// </summary>
        /// <param name="state"></param>
        protected override void RemoveFromCrm(SyncState state)
        {
            if (state.CrmType == MeetingsSynchroniser.CrmModule)
            {
                this.AddOrUpdateItemFromOutlookToCrm((SyncState<Outlook.AppointmentItem>)state);
            }
            else
            {
                base.RemoveFromCrm(state);
            }
        }


        /// <summary>
        /// Typically, when handling an item missing from outlook, the outlook item is missing and so can't
        /// be relied on; treat this record as representing the current, known state of the item.
        /// </summary>
        /// <param name="record">A record fetched from CRM representing the current state of the item.</param>
        /// <param name="syncState">The sync state representing the item.</param>
        /// <param name="crmModule">The name/key of the CRM module in which the item exists.</param>
        private void HandleItemMissingFromOutlook(EntryValue record, SyncState<Outlook.AppointmentItem> syncState, string crmModule)
        {
            try
            {
                if (record.GetValueAsDateTime("date_start") > DateTime.Now && crmModule == MeetingsSynchroniser.CrmModule)
                {
                    /* meeting in the future: mark it as canceled, do not delete it */
                    record.GetBinding("status").value = "NotHeld";

                    string description = record.GetValue("description").ToString();
                    if (!description.StartsWith(AppointmentsSynchroniser<SyncStateType>.CanceledPrefix))
                    {
                        record.GetBinding("description").value = $"{AppointmentsSynchroniser<SyncStateType>.CanceledPrefix}: {description}";
                        RestAPIWrapper.SetEntry(record.nameValueList, crmModule);
                    }
                }
                else
                {
                    /* meeting in the past: just delete it */
                    this.RemoveFromCrm(syncState);
                    this.RemoveItemSyncState(syncState);
                }
            }
            catch (Exception any)
            {
                /* what could possibly go wrong? */
                this.Log.Error($"Failed in HandleItemMissingFromOutlook for CRM Id {syncState.CrmEntryId}", any);
            }
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
            RemoveSynchronisationPropertyFromOutlookItem(olItem, SyncStateManager.CrmIdPropertyName);
            RemoveSynchronisationPropertyFromOutlookItem(olItem, SyncStateManager.TypePropertyName);
            RemoveSynchronisationPropertyFromOutlookItem(olItem, SyncStateManager.ModifiedDatePropertyName);
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
                catch (Exception any)
                {
                    Globals.ThisAddIn.Log.Warn($"Unexpected error in RemoveSynchronisationPropertyFromOutlookItem", any);
                }
                finally
                {
                    olItem.Save();
                }
            }
        }


        /// <summary>
        /// Set the meeting status of this `olItem` from this `crmItem`.
        /// </summary>
        /// <param name="olItem">The Outlook item to update.</param>
        /// <param name="crmItem">The CRM item to use as source.</param>
        protected abstract void SetMeetingStatus(Outlook.AppointmentItem olItem, EntryValue crmItem);

        /// <summary>
        /// We should delete an item from CRM if it already exists in CRM, but it is now private.
        /// </summary>
        /// <param name="olItem">The Outlook item</param>
        /// <returns>true if the Outlook item should be deleted from CRM.</returns>
        private bool ShouldDeleteFromCrm(Outlook.AppointmentItem olItem)
        {
            Outlook.UserProperty olPropertyEntryId = olItem.UserProperties[SyncStateManager.CrmIdPropertyName];
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
            var syncConfigured = SyncDirection.AllowOutbound(Direction);
            string organiser = olItem.Organizer;
            var currentUser = Application.Session.CurrentUser;
            var exchangeUser = currentUser.AddressEntry.GetExchangeUser();
            var currentUserName = exchangeUser == null ? 
                Application.Session.CurrentUser.Name:
                exchangeUser.Name;
            string crmId = olItem.UserProperties[SyncStateManager.CrmIdPropertyName]?.Value;

            return olItem != null &&
                syncConfigured && 
                olItem.Sensitivity == Outlook.OlSensitivity.olNormal &&
                /* If there is a valid crmId it's arrived via CRM and is therefore safe to save to CRM;
                 * if the current user is the organiser, AND there's no valid CRM id, then it's a new one
                 * that the current user made, and we should save it to CRM. */
                (!string.IsNullOrEmpty(crmId) || currentUserName == organiser) &&
                /* Microsoft Conferencing Add-in creates temporary items with names which start 
                 * 'PLEASE IGNORE' - we should not sync these. */
                !olItem.Subject.StartsWith(MSConfTmpSubjectPrefix);
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
            Log.Debug($"{this.GetType().Name}.SyncFolder: '{crmModule}'");

            try
            {
                /* this.ItemsSyncState already contains items to be synced. */
                var untouched = new HashSet<SyncState<Outlook.AppointmentItem>>(SyncStateManager.Instance.GetSynchronisedItems<SyncStateType>());
                IList<EntryValue> records = MergeRecordsFromCrm(folder, crmModule, untouched);

                AddOrUpdateItemsFromCrmToOutlook(records, folder, untouched, crmModule);

                EntryValue[] invited = RestAPIWrapper.GetRelationships("Users",
                    RestAPIWrapper.GetUserId(), crmModule.ToLower(),
                    RestAPIWrapper.GetSugarFields(crmModule));
                if (invited != null)
                {
                    AddOrUpdateItemsFromCrmToOutlook(invited, folder, untouched, crmModule);
                }

                try
                {
                    this.ResolveUnmatchedItems(untouched);
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
                Outlook.UserProperty olPropertyModifiedDate = olItem.UserProperties[SyncStateManager.ModifiedDatePropertyName];

                if (olPropertyModifiedDate.Value != crmItem.GetValueAsString("date_modified"))
                {
                    try
                    {
                        olItem.Subject = crmItem.GetValueAsString("name");
                        olItem.Body = crmItem.GetValueAsString("description");
                        if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("date_start")))
                        {
                            UpdateOutlookDetails(crmType, crmItem, date_start, olItem);
                        }

                        EnsureSynchronisationPropertiesForOutlookItem(olItem, crmItem, crmType);
                        LogItemAction(syncState.OutlookItem, "AppointmentSyncing.UpdateExistingOutlookItemFromCrm, item saved");
                    }
                    catch (Exception any)
                    {
                        Globals.ThisAddIn.Log.Warn($"Unexpected error in UpdateExistingOutlookItemFromCrm", any);
                    }
                    finally
                    {
                        this.SaveItem(olItem);
                    }
                }
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
        protected virtual void UpdateOutlookDetails(string crmType, EntryValue crmItem, DateTime date_start, Outlook.AppointmentItem olItem)
        {
            try
            {
                olItem.Start = date_start;
                var minutesString = crmItem.GetValueAsString("duration_minutes");
                var hoursString = crmItem.GetValueAsString("duration_hours");

                int minutes = string.IsNullOrWhiteSpace(minutesString) ? 0 : int.Parse(minutesString);
                int hours = string.IsNullOrWhiteSpace(hoursString) ? 0 : int.Parse(hoursString);

                olItem.Duration = minutes + hours * 60;
            }

            finally
            {
                this.SaveItem(olItem);
            }
        }


        internal override string GetOutlookEntryId(Outlook.AppointmentItem olItem)
        {
            return olItem.EntryID;
        }

        protected override string GetCrmEntryId(Outlook.AppointmentItem olItem)
        {
            return olItem?.UserProperties[SyncStateManager.CrmIdPropertyName]?.Value.ToString();
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
        /// Add an address resolution composed from this module name and record to the cache.
        /// </summary>
        /// <param name="moduleName">The name of the module in which the record was found</param>
        /// <param name="record">The record.</param>
        protected void CacheAddressResolutionData(string moduleName, LinkRecord record)
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
        protected void CacheAddressResolutionData(AddressResolutionData resolution)
        {
            List<AddressResolutionData> resolutions;

            if (this.meetingRecipientsCache.ContainsKey(resolution.emailAddress))
            {
                resolutions = this.meetingRecipientsCache[resolution.emailAddress];
            }
            else
            {
                resolutions = new List<AddressResolutionData>();
                this.meetingRecipientsCache[resolution.emailAddress] = resolutions;
            }

            if (!resolutions.Any(x => x.moduleId == resolution.moduleId && x.moduleName == resolution.moduleName))
            {
                resolutions.Add(resolution);
            }

            Log.Debug($"Successfully cached recipient {resolution.emailAddress} => {resolution.moduleName}, {resolution.moduleId}.");
        }


        protected void CacheAddressResolutionData(EntryValue crmItem)
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


        /// <summary>
        /// Used for caching data for resolving email addresses to CRM records.
        /// </summary>
        protected class AddressResolutionData
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

            public AddressResolutionData(string moduleName, Dictionary<string, object> data)
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
