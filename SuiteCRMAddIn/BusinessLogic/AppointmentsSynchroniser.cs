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


#region

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using SuiteCRMAddIn.Exceptions;
using SuiteCRMAddIn.Extensions;
using SuiteCRMAddIn.Properties;
using SuiteCRMAddIn.ProtoItems;
using SuiteCRMClient;
using SuiteCRMClient.Logging;
using SuiteCRMClient.RESTObjects;
using Exception = System.Exception;

#endregion

namespace SuiteCRMAddIn.BusinessLogic
{
    /// <summary>
    ///     Handles the synchronisation of appointments between Outlook and CMS.
    /// </summary>
    public abstract class AppointmentsSynchroniser<SyncStateType> : Synchroniser<AppointmentItem, SyncStateType>
        where SyncStateType : SyncState<AppointmentItem>
    {
        /// <summary>
        ///     The name of the organiser synchronisation property
        /// </summary>
        public const string OrganiserPropertyName = "SOrganiser";

        /// <summary>
        ///     Microsoft Conferencing Add-in creates temporary items whose names begin
        ///     'PLEASE IGNORE'. We should not sync these.
        /// </summary>
        public const string MSConfTmpSubjectPrefix = "PLEASE IGNORE";

        /// <summary>
        ///     Prefix for meetings which have been canceled.
        /// </summary>
        private const string CanceledPrefix = "CANCELED";

        /// <summary>
        ///     A cache of email addresses to CRM modules and identities
        /// </summary>
        protected Dictionary<string, List<AddressResolutionData>> meetingRecipientsCache =
            new Dictionary<string, List<AddressResolutionData>>();


        public AppointmentsSynchroniser(string name, SyncContext context)
            : base(name, context)
        {
            // TODO: Also need to fetch appointments to which the current user is invited.
            fetchQueryPrefix = new StringBuilder("assigned_user_id = '{0}'")
                .Append($" and date_start > '{string.Format("{0:yyyy-MM-dd HH:mm:ss}", GetStartDate())}' ").ToString();
        }

        /// <summary>
        ///     Get the id of the record with the specified `smtpAddress` in the module with the specified `moduleName`.
        /// </summary>
        /// <param name="smtpAddress">The SMTP email address to be sought.</param>
        /// <param name="moduleName">The name of the module in which to seek it.</param>
        /// <returns>The corresponding id, if present, else the empty string.</returns>
        public CrmId GetInviteeIdBySmtpAddress(string smtpAddress, string moduleName)
        {
            var bob = new StringBuilder($"({moduleName.ToLower()}.id in ")
                .Append($"(select eabr.bean_id from email_addr_bean_rel eabr ")
                .Append($"INNER JOIN email_addresses ea on eabr.email_address_id = ea.id ")
                .Append($"where eabr.bean_module = '{moduleName}' ")
                .Append($"and ea.email_address LIKE '%{RestAPIWrapper.MySqlEscape(smtpAddress)}%'))");

            var query = bob.ToString();

            Log.Debug($"AppointmentSyncing.GetID: query = `{query}`");

            string[] fields = {"id"};
            var entries = RestAPIWrapper.GetEntryList(moduleName, query, Settings.Default.SyncMaxRecords,
                "date_entered DESC", 0, false, fields);

            return entries.result_count > 0
                ? CrmId.Get(RestAPIWrapper.GetValueByKey(entries.entry_list[0], "id"))
                : CrmId.Empty;
        }

        public override MAPIFolder GetDefaultFolder()
        {
            return Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
        }

        /// <summary>
        ///     #2246: Discriminate between calls and meetings when adding and updating.
        /// </summary>
        protected override void OutlookItemAdded(AppointmentItem olItem)
        {
            if (Globals.ThisAddIn.IsLicensed)
                lock (enqueueingLock)
                {
                    try
                    {
                        var crmId = olItem.GetCrmId();
                        var entryId = olItem.EntryID;

                        if (CrmId.IsValid(crmId))
                        {
                            var existing =
                                SyncStateManager.Instance.GetExistingSyncState(string.Empty, crmId) as AppointmentSyncState;

                            /* we have an existing item with the same CRM id: suspicious */
                            var fromVcal = existing.OutlookItem.GetVCalId();

                            if (existing != null && !string.IsNullOrEmpty(fromVcal) && fromVcal != crmId.ToString())
                            {
                                /* OK, its GlobalAppointmentId is wrong, it must have come via sync. Delete it. */
                                Log.Debug(
                                    $"Apparent case of item synced from CRM and then received via email/vCal. CRM id is {crmId}. Deleting bad copy.");
                                RemoveItemAndSyncState(existing);
                            }
                        }

                        Log.Debug($"OutlookItemAdded: entry, CRM id = {crmId}; Outlook ID = {entryId}");
                        if (olItem.IsCall())
                            base.OutlookItemAdded(olItem, Globals.ThisAddIn.CallsSynchroniser);
                        else
                            base.OutlookItemAdded(olItem, Globals.ThisAddIn.MeetingsSynchroniser);
                        Log.Debug($"OutlookItemAdded: exit, CRM id = {crmId}; Outlook ID = {entryId}");
                    }
                    finally
                    {
                        SaveItem(olItem);
                    }
                }
            else
                Log.Warn(
                    $"Synchroniser.OutlookItemAdded: item {GetOutlookEntryId(olItem)} not added because not licensed");
        }

        /// <summary>
        ///     #2246: Discriminate between calls and meetings when adding and updating.
        /// </summary>
        protected override void OutlookItemChanged(AppointmentItem olItem)
        {
            if (Globals.ThisAddIn.IsLicensed)
                try
                {
                    var crmId = olItem.GetCrmId();
                    Log.Debug($"OutlookItemChanged: entry, CRM id = {crmId}; Outlook ID = {olItem.EntryID}");

                    if (olItem.IsCall())
                        base.OutlookItemChanged(olItem, Globals.ThisAddIn.CallsSynchroniser);
                    else
                        base.OutlookItemChanged(olItem, Globals.ThisAddIn.MeetingsSynchroniser);

                    Log.Debug($"OutlookItemChanged: exit, CRM id = {crmId}; Outlook ID = {olItem.EntryID}");
                }
                catch (BadStateTransition bst)
                {
                    if (bst.To != TransmissionState.Transmitted)
                        Log.Warn("Bad state transition in OutlookItemChanged", bst);
                }
                finally
                {
                    SaveItem(olItem);
                }
            else
                Log.Warn(
                    $"Synchroniser.OutlookItemAdded: item {GetOutlookEntryId(olItem)} not updated because not licensed");
        }

        protected override void SaveItem(AppointmentItem olItem)
        {
            try
            {
                if (olItem != null && olItem.IsValid())
                {
                    olItem?.Save();
                    try
                    {
                        LogItemAction(olItem, "AppointmentSyncing.SaveItem, saved item");
                    }
                    catch (InvalidCrmIdException)
                    {
                        Log.Debug(
                            $"AppointmentSyncing.SaveItem, saved item '{olItem.Subject}' {olItem.EntryID} (no valid CRM id)");
                    }
                }
            }
            catch (Exception any)
            {
                try
                {
                    ErrorHandler.Handle($"Failure while saving appointment {olItem?.Subject}", any);
                }
                catch (COMException comx)
                {
                    ErrorHandler.Handle(
                        "Failure while trying to save appointment, appointment has probably been deleted.", comx);
                }
            }
        }


        /// <summary>
        ///     Ensure that this Outlook item has a property of this name with this value.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        protected override void EnsureSynchronisationPropertyForOutlookItem(AppointmentItem olItem, string name,
            string value)
        {
            try
            {
                var olProperty = olItem.UserProperties[name] ??
                                 olItem.UserProperties.Add(name, OlUserPropertyType.olText);

                if (!olProperty.Value.Equals(value))
                    try
                    {
                        if (string.IsNullOrEmpty(value))
                        {
                            olProperty.Delete();
                        }
                        else
                        {
                            olProperty.Value = value;
                            Log.Debug(
                                $"AppointmentSyncing.EnsureSynchronisationPropertyForOutlookItem: Set property {name} to value {value} on item {olItem.Subject}");
                        }
                    }
                    catch (Exception any)
                    {
                        ErrorHandler.Handle($"Failed to set property {name} to value {value} on item {olItem.Subject}",
                            any);
                    }
                    finally
                    {
                        SaveItem(olItem);
                    }
            }
            catch (Exception any)
            {
                ErrorHandler.Handle($"Failed to set property {name} to value {value} on item {olItem.Subject}", any);
            }
        }


        /// <summary>
        ///     If a meeting was created in another Outlook we should NOT sync it with CRM because if we do we'll create
        ///     duplicates. Only the Outlook which created it should sync it.
        /// </summary>
        /// <param name="folder">The folder to synchronise into.</param>
        /// <param name="crmType">The CRM type of the candidate item.</param>
        /// <param name="crmItem">The candidate item from CRM.</param>
        /// <returns>True if it's offered to us by CRM with its Outlook ID already populated.</returns>
        protected override bool ShouldAddOrUpdateItemFromCrmToOutlook(MAPIFolder folder, string crmType,
            EntryValue crmItem)
        {
            var outlookId = crmItem.GetValueAsString("outlook_id");
            /* we're good if it's a meeting... */
            var result = crmType == DefaultCrmModule;
            /* provided it doesn't already have an Outlook id */
            result &= string.IsNullOrWhiteSpace(outlookId);
            /* and we're also good if we've already got it */
            result |= SyncStateManager.Instance.GetExistingSyncState(crmItem) != null;

            if (!result)
                Log.Debug(
                    $"ShouldAddOrUpdateItemFromCrmToOutlook: not syncing meeting `{crmItem.GetValueAsString("name")}` as it appears to originate from another Outlook instance.");

            return result;
        }


        /// <summary>
        ///     Add an item existing in CRM but not found in Outlook to Outlook.
        /// </summary>
        /// <remarks>
        ///     This method is disconcertingly different from equivalent methods in other synchronisers;
        ///     TODO: the differences ought to be thought about.
        /// </remarks>
        /// <see cref="ContactSynchroniser.AddNewItemFromCrmToOutlook(Microsoft.Office.Interop.Outlook.MAPIFolder, EntryValue)" />
        /// <param name="appointmentsFolder">The Outlook folder in which the item should be stored.</param>
        /// <param name="crmType">The CRM type of the item from which values are to be taken.</param>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <param name="dateStart">The state date/time of the item, adjusted for timezone.</param>
        /// <returns>A sync state object for the new item.</returns>
        protected virtual SyncStateType AddNewItemFromCrmToOutlook(
            MAPIFolder appointmentsFolder,
            string crmType,
            EntryValue crmItem,
            DateTime dateStart)
        {
            SyncStateType newState = null;
            AppointmentItem olItem = null;

            Log.Debug(
                string.Format(
                    $"{GetType().Name}.AddNewItemFromCrmToOutlook, entry id is '{crmItem.GetValueAsString("id")}', creating in Outlook."));

            /*
             * There's a nasty little bug (#223) where Outlook offers us back in a different thread
             * the item we're creating, before we're able to set up the sync state which marks it
             * as already known. By locking on the enqueueing lock here, we should prevent that.
             */
            lock (enqueueingLock)
            {
                try
                {
                    var crmId = crmItem.GetValueAsString("id");

                    olItem = appointmentsFolder.Items.Add(OlItemType.olAppointmentItem);

                    olItem.Subject = crmItem.GetValueAsString("name");
                    olItem.Body = crmItem.GetValueAsString("description");
                    SetMeetingStatus(olItem, crmItem);

                    /* set the SEntryID property quickly, create the sync state and save the item, to reduce howlaround */
                    EnsureSynchronisationPropertiesForOutlookItem(olItem, crmItem, crmType);

                    LogItemAction(olItem, "AppointmentSyncing.AddNewItemFromCrmToOutlook");
                    if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("date_start")))
                    {
                        olItem.Start = dateStart;
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
                        try
                        {
                            olItem.Save();
                            newState = SyncStateManager.Instance.GetOrCreateSyncState(olItem) as SyncStateType;
                            newState.SetNewFromCRM();
                        }
                        catch (COMException comx)
                        {
                            // this really, really shouldn't happen!
                            ErrorHandler.Handle(comx);
                        }
                    }
                }
            }

            return newState;
        }


        /// <summary>
        ///     Set this outlook item's duration, but also end time and location, from this CRM item.
        /// </summary>
        /// <param name="crmType">The type of the CRM item.</param>
        /// <param name="crmItem">The CRM item.</param>
        /// <param name="olItem">The Outlook item.</param>
        protected void SetOutlookItemDuration(string crmType, EntryValue crmItem, AppointmentItem olItem)
        {
            try
            {
                SetOutlookItemDuration(crmItem, olItem);
            }
            catch (Exception any)
            {
                ErrorHandler.Handle("Failed while setting Outlook item duration", any);
            }
            finally
            {
                SaveItem(olItem);
            }
        }

        /// <summary>
        ///     Set this outlook item's duration from this CRM item.
        /// </summary>
        /// <param name="crmItem">The CRM item.</param>
        /// <param name="olItem">The Outlook item.</param>
        protected virtual void SetOutlookItemDuration(EntryValue crmItem, AppointmentItem olItem)
        {
            int minutes = 0, hours = 0;

            if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("duration_minutes")))
                minutes = int.Parse(crmItem.GetValueAsString("duration_minutes"));
            if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("duration_hours")))
                hours = int.Parse(crmItem.GetValueAsString("duration_hours"));

            var durationMinutes = minutes + hours * 60;

            olItem.Duration = durationMinutes;
        }

        /// <summary>
        ///     Specialisation: in addition to the standard properties, meetings also require an organiser property.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmItem">The CRM item.</param>
        /// <param name="type">The value for the SType property (CRM module name).</param>
        protected override void EnsureSynchronisationPropertiesForOutlookItem(AppointmentItem olItem,
            EntryValue crmItem, string type)
        {
            base.EnsureSynchronisationPropertiesForOutlookItem(olItem, crmItem, type);
            if (DefaultCrmModule.Equals(type))
                EnsureSynchronisationPropertyForOutlookItem(olItem, OrganiserPropertyName,
                    crmItem.GetValueAsString("assigned_user_id"));
        }

        /// <summary>
        ///     Add the item implied by this SyncState, which may not exist in CRM, to CRM.
        /// </summary>
        /// <param name="syncState">The sync state.</param>
        /// <returns>The id of the entry added or updated.</returns>
        internal override CrmId AddOrUpdateItemFromOutlookToCrm(SyncState<AppointmentItem> syncState)
        {
            var olItem = syncState.OutlookItem;

            lock (enqueueingLock)
            {
                try
                {
                    var result = CrmId.Empty;

                    if (olItem == null || !olItem.IsValid())
                    {
                        HandleItemMissingFromOutlook(syncState);
                    }
                    else if (ShouldAddOrUpdateItemFromOutlookToCrm(olItem))
                        if (ShouldDeleteFromCrm(olItem))
                        {
                            LogItemAction(olItem, "AppointmentSyncing.AddOrUpdateItemFromOutlookToCrm: Deleting");

                            DeleteFromCrm(olItem);
                        }
                        else if (ShouldDespatchToCrm(olItem))
                        {
                            result = base.AddOrUpdateItemFromOutlookToCrm(syncState);

                            if (CrmId.IsValid(result) && CrmId.IsValid(olItem.GetCrmId()))
                            {
                                if (syncState is CallSyncState)
                                {
                                    SetCrmRelationshipFromOutlook(Globals.ThisAddIn.CallsSynchroniser, result, "Users",
                                        CrmId.Get(RestAPIWrapper.GetUserId()));
                                }
                                else
                                {
                                    SetCrmRelationshipFromOutlook(Globals.ThisAddIn.MeetingsSynchroniser, result, "Users",
                                        CrmId.Get(RestAPIWrapper.GetUserId()));
                                }
                            }
                        }
                        else
                        {
                            LogItemAction(olItem, "AppointmentSyncing.AddItemFromOutlookToCrm, Not despatching");
                        }
                    else
                        LogItemAction(olItem, "AppointmentSyncing.AddItemFromOutlookToCrm, Not enabled");

                    return result;
                }
                catch (COMException)
                {
                    HandleItemMissingFromOutlook(syncState);
                    return syncState.CrmEntryId;
                }
            }
        }


        /// <summary>
        /// Construct a JSON packet representing the Outlook item of this sync state, and despatch 
        /// it to CRM.
        /// </summary>
        /// <param name="syncState">The sync state.</param>
        /// <returns>The CRM id of the object created or modified.</returns>
        protected override CrmId ConstructAndDespatchCrmItem(SyncState<AppointmentItem> syncState)
        {
            return CrmId.Get(RestAPIWrapper.SetEntry(new ProtoAppointment<SyncStateType>(syncState.OutlookItem).AsNameValues(),
                DefaultCrmModule));
        }


        /// <summary>
        ///     Delete this Outlook item from CRM, and tidy up afterwards.
        /// </summary>
        /// <param name="olItem">The Outlook item to delete.</param>
        private void DeleteFromCrm(AppointmentItem olItem)
        {
            if (olItem != null)
            {
                /* Remove the magic properties */
                RemoveSynchronisationPropertiesFromOutlookItem(olItem);
                var syncStateForItem = SyncStateManager.Instance.GetExistingSyncState(olItem);
                if (syncStateForItem != null)
                {
                    RemoveFromCrm(syncStateForItem);
                    RemoveItemSyncState(syncStateForItem);
                }
            }
        }

        /// <summary>
        ///     Get all items in this appointments folder. Should be called just once (per folder?)
        ///     when the add-in starts up; initialises the SyncState list.
        /// </summary>
        /// <param name="appointmentsFolder">The folder to scan.</param>
        protected override void LinkOutlookItems(MAPIFolder appointmentsFolder)
        {
            try
            {
                var deletionCandidates = new List<AppointmentItem>();
                foreach (AppointmentItem olItem in appointmentsFolder.Items)
                {
                    if (olItem.IsValid())
                    {
                        try
                        {
                            if (olItem.Start >= GetStartDate())
                            {
                                var olPropertyModified = olItem.UserProperties[SyncStateManager.ModifiedDatePropertyName];
                                var olPropertyType = olItem.UserProperties[SyncStateManager.TypePropertyName];
                                var olPropertyEntryId = olItem.GetCrmId();
                                if (olPropertyModified != null &&
                                    olPropertyType != null &&
                                    olPropertyEntryId != null)
                                    LogItemAction(olItem,
                                        "AppointmentSyncing.LinkOutlookItems: Adding known item to queue");
                                else
                                    LogItemAction(olItem,
                                        "AppointmentSyncing.LinkOutlookItems: Adding unknown item to queue");

                                SyncStateManager.Instance.GetOrCreateSyncState(olItem).SetPresentAtStartup();
                            }
                        }
                        catch (ProbableDuplicateItemException<AppointmentItem>)
                        {
                            deletionCandidates.Add(olItem);
                        }
                    }
                }

                foreach (var toDelete in deletionCandidates)
                    toDelete.Delete();
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle($"Failed while trying to index {DefaultCrmModule}", ex);
            }
        }

        /// <summary>
        ///     Log a message regarding this Outlook appointment.
        /// </summary>
        /// <param name="olItem">The outlook item.</param>
        /// <param name="message">The message to be logged.</param>
        internal override void LogItemAction(AppointmentItem olItem, string message)
        {
            try
            {
                if (olItem != null && olItem.IsValid())
                {
                    var crmId = IsEnabled() ? olItem.GetCrmId() : CrmId.Empty;
                    if (CrmId.IsInvalid(crmId)) crmId = CrmId.Empty;

                    var bob = new StringBuilder();
                    bob.Append($"{message}:\n\tOutlook Id  : {olItem.EntryID}")
                        .Append($"\n\tGlobal Id   : {olItem.GlobalAppointmentID}")
                        .Append(this.IsEnabled() ? $"\n\tCRM Id      : {crmId}" : string.Empty)
                        .Append($"\n\tSubject     : '{olItem.Subject}'")
                        .Append($"\n\tSensitivity : {olItem.Sensitivity}")
                        .Append($"\n\tStatus      : {olItem.MeetingStatus}")
                        .Append($"\n\tReminder set: {olItem.ReminderSet}")
                        .Append($"\n\tOrganiser   : {olItem.Organizer}")
                        .Append($"\n\tOutlook User: {Globals.ThisAddIn.Application.GetCurrentUsername()}")
                        .Append($"\n\tTxState     : {SyncStateManager.Instance.GetExistingSyncState(olItem)?.TxState}")
                        .Append($"\n\tRecipients  :\n");

                    foreach (Recipient recipient in olItem.Recipients)
                        bob.Append(
                            $"\t\t{recipient.Name}: {recipient.GetSmtpAddress()} - ({recipient.MeetingResponseStatus})\n");
                    Log.Info(bob.ToString());
                }
            }
            catch (Exception e)
            {
                Log.Error($"Unexpected error while trying to log action '{message}'", e);
            }
        }


        /// <summary>
        ///     Update a single appointment in the specified Outlook folder with changes from CRM, but
        ///     only if its start date is fewer than five days in the past.
        /// </summary>
        /// <param name="folder">The folder to synchronise into.</param>
        /// <param name="crmType">The CRM type of the candidate item.</param>
        /// <param name="crmItem">The candidate item from CRM.</param>
        /// <returns>The synchronisation state of the item updated (if it was updated).</returns>
        protected override SyncState<AppointmentItem> AddOrUpdateItemFromCrmToOutlook(
            MAPIFolder folder,
            string crmType,
            EntryValue crmItem)
        {
            var existing = SyncStateManager.Instance.GetExistingSyncState(crmItem);
            var result = existing as SyncStateType;

            var dateStart = crmItem.GetValueAsDateTime("date_start");

            if (dateStart >= GetStartDate())
            {
                /* search for the item among the sync states I already know about */
                if (existing == null)
                {
                    /* check for howlaround */
                    var matches = FindMatches(crmItem);

                    if (!matches.Any())
                    {
                        /* didn't find it, so add it to Outlook */
                        result = AddNewItemFromCrmToOutlook(folder, crmType, crmItem, dateStart);
                    }
                    else
                    {
                        var withoutCrmId = matches.Where(x => CrmId.IsInvalid(x.CrmEntryId)).ToList();
                        var crmId = CrmId.Get(crmItem.id);
                        if (withoutCrmId.Any())
                        {
                            result = withoutCrmId.ElementAt(0) as SyncStateType;
                            if (result != null)
                            {
                                result.CrmEntryId = crmId;
                                result.OutlookItem.SetCrmId(crmId);
                                UpdateExistingOutlookItemFromCrm(crmType, crmItem, dateStart, result);
                            }
                        }
                        else
                        {
                            result = matches.ElementAt(0) as SyncStateType;
                            Log.Warn(
                                $"Howlaround detected? Appointment '{crmItem.GetValueAsString("name")}' offered with id {crmId}, expected {matches[0].CrmEntryId}, {matches.Count} duplicates");
                        }
                    }
                }
                else if (result != null)
                {
                    /* found it, so update it from the CRM item */
                    UpdateExistingOutlookItemFromCrm(crmType, crmItem, dateStart, result);
                }
                else
                {
                    throw new UnexpectedSyncStateClassException($"{GetType().Name}", existing);
                }

                existing?.SaveItem();
            }

            return result;
        }

        //internal override void HandleItemMissingFromOutlook(SyncState<AppointmentItem> syncState)
        //{
        //    if (syncState.CrmType == MeetingsSynchroniser.CrmModule)
        //    {
        //        /* typically, when this method is called, the Outlook Item will already be invalid, and if it is not,
        //         * it may become invalid during the execution of this method. So this method CANNOT depend on any
        //         * values taken from the Outlook item. */
        //        var entries = RestAPIWrapper.GetEntryList(
        //            syncState.CrmType, $"id = {syncState.CrmEntryId}",
        //            Settings.Default.SyncMaxRecords,
        //            "date_entered DESC", 0, false, null);

        //        if (entries.entry_list.Any())
        //            HandleItemMissingFromOutlook(entries.entry_list[0], syncState, syncState.CrmType);
        //    }
        //}


        /// <summary>
        ///     Sets up a CRM relationship to mimic an Outlook relationship
        /// </summary>
        /// <param name="sync">The synchroniser tp forward to.</param>
        /// <param name="meetingId">The ID of the appointment.</param>
        /// <param name="recipient">The outlook recipient representing the person to link with.</param>
        /// <param name="foreignModule">the name of the module we're seeking to link with.</param>
        /// <returns>True if a relationship was created.</returns>
        protected CrmId SetCrmRelationshipFromOutlook<T, S>(Synchroniser<T, S> sync, CrmId meetingId,
            Recipient recipient, string foreignModule)
            where T : class
            where S : SyncState<T>
        {
            var foreignId = GetInviteeIdBySmtpAddress(recipient.GetSmtpAddress(), foreignModule);

            return CrmId.IsValid(foreignId) &&
                   SetCrmRelationshipFromOutlook(sync, meetingId, foreignModule, foreignId)
                ? foreignId
                : CrmId.Empty;
        }


        /// <summary>
        ///     Sets up a CRM relationship to mimic an Outlook relationship
        /// </summary>
        /// <param name="sync">the synchroniser to despatch to</param>
        /// <param name="meetingId">The meeting id.</param>
        /// <param name="resolution">Address resolution data from the cache.</param>
        /// <returns>True if a relationship was created.</returns>
        protected bool SetCrmRelationshipFromOutlook<T, S>(Synchroniser<T, S> sync, CrmId meetingId,
            AddressResolutionData resolution)
            where T : class
            where S : SyncState<T>
        {
            return SetCrmRelationshipFromOutlook(sync, meetingId, resolution.ModuleName, resolution.ModuleId);
        }


        /// <summary>
        ///     Sets up a CRM relationship to mimic an Outlook relationship
        /// </summary>
        /// <param name="sync">the synchroniser to despatch to.</param>
        /// <param name="meetingId">The ID of the meeting</param>
        /// <param name="foreignModule">the name of the module we're seeking to link with.</param>
        /// <param name="foreignId">The id in the foreign module of the record we're linking to.</param>
        /// <returns>True if a relationship was created.</returns>
        protected bool SetCrmRelationshipFromOutlook<T, S>(Synchroniser<T, S> sync, CrmId meetingId,
            string foreignModule, CrmId foreignId)
            where T : class
            where S : SyncState<T>
        {
            return CrmId.IsValid(foreignId) &&
                   RestAPIWrapper.SetRelationshipUnsafe(new SetRelationshipParams
                   {
                       module2 = sync.DefaultCrmModule,
                       module2_id = meetingId.ToString(),
                       module1 = foreignModule,
                       module1_id = foreignId.ToString()
                   });
        }


        /// <summary>
        ///     Typically, when handling an item missing from outlook, the outlook item is missing and so can't
        ///     be relied on; treat this record as representing the current, known state of the item.
        /// </summary>
        /// <param name="record">A record fetched from CRM representing the current state of the item.</param>
        /// <param name="syncState">The sync state representing the item.</param>
        /// <param name="crmModule">The name/key of the CRM module in which the item exists.</param>
        private void HandleItemMissingFromOutlook(EntryValue record, SyncState<AppointmentItem> syncState,
            string crmModule)
        {
            try
            {
                if (record.GetValueAsDateTime("date_start") > DateTime.Now &&
                    crmModule == MeetingsSynchroniser.CrmModule)
                {
                    /* meeting in the future: mark it as canceled, do not delete it */
                    record.GetBinding("status").value = "NotHeld";

                    var description = record.GetValue("description").ToString();
                    if (!description.StartsWith(CanceledPrefix))
                    {
                        record.GetBinding("description").value = $"{CanceledPrefix}: {description}";
                        RestAPIWrapper.SetEntry(record.nameValueList, crmModule);
                    }
                }
                else
                {
                    /* meeting in the past: just delete it */
                    RemoveFromCrm(syncState);
                    RemoveItemSyncState(syncState);
                }
            }
            catch (Exception any)
            {
                /* what could possibly go wrong? */
                ErrorHandler.Handle(
                    $"Failed while attempting to handle item missing from Outlook; CRM Id is {syncState.CrmEntryId}",
                    any);
            }
        }

        protected override bool IsMatch(AppointmentItem olItem, EntryValue crmItem)
        {
            bool result;
            try
            {
                result = olItem.Subject == crmItem.GetValueAsString("name") &&
                       crmItem.GetValueAsDateTime("date_start") == olItem.Start;
            }
            catch (COMException)
            {
                result = false;
            }

            return result;
        }

        /// <summary>
        ///     Remove the synchronisation properties from this Outlook item.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        private static void RemoveSynchronisationPropertiesFromOutlookItem(AppointmentItem olItem)
        {
            RemoveSynchronisationPropertyFromOutlookItem(olItem, SyncStateManager.CrmIdPropertyName);
            RemoveSynchronisationPropertyFromOutlookItem(olItem, SyncStateManager.TypePropertyName);
            RemoveSynchronisationPropertyFromOutlookItem(olItem, SyncStateManager.ModifiedDatePropertyName);
        }

        /// <summary>
        ///     Ensure that this Outlook item does not have a property of this name.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="name">The name.</param>
        private static void RemoveSynchronisationPropertyFromOutlookItem(AppointmentItem olItem, string name)
        {
            var found = 0;
            /* typical Microsoft, you can only remove a user property by its 1-based number */

            for (var i = 1; i <= olItem.UserProperties.Count; i++)
                if (olItem.UserProperties[i].Name == name)
                {
                    found = i;
                    break;
                }

            if (found > 0)
                try
                {
                    olItem.UserProperties.Remove(found);
                }
                catch (Exception any)
                {
                    Globals.ThisAddIn.Log.Warn($"Unexpected error in RemoveSynchronisationPropertyFromOutlookItem",
                        any);
                }
                finally
                {
                    olItem.Save();
                }
        }


        /// <summary>
        ///     Set the meeting status of this `olItem` from this `crmItem`.
        /// </summary>
        /// <param name="olItem">The Outlook item to update.</param>
        /// <param name="crmItem">The CRM item to use as source.</param>
        protected abstract void SetMeetingStatus(AppointmentItem olItem, EntryValue crmItem);

        /// <summary>
        ///     We should delete an item from CRM if it already exists in CRM, but it is now private.
        /// </summary>
        /// <param name="olItem">The Outlook item</param>
        /// <returns>true if the Outlook item should be deleted from CRM.</returns>
        private bool ShouldDeleteFromCrm(AppointmentItem olItem)
        {
            var result = CrmId.IsValid(olItem.GetCrmId()) && olItem.Sensitivity != OlSensitivity.olNormal;

            LogItemAction(olItem, $"ShouldDeleteFromCrm returning {result}");

            return result;
        }

        /// <summary>
        ///     True if we should despatch this item to CRM, else false.
        /// </summary>
        /// <param name="olItem"></param>
        /// <returns>true iff settings.SyncCalendar is true, the item is not null, and it is not private (normal sensitivity)</returns>
        private bool ShouldDespatchToCrm(AppointmentItem olItem)
        {
            var syncConfigured = SyncDirection.AllowOutbound(Direction);
            var organiser = olItem.Organizer;
            var currentUser = Application.Session.CurrentUser;
            var exchangeUser = currentUser.AddressEntry.GetExchangeUser();
            var currentUserName = exchangeUser == null ? Application.Session.CurrentUser.Name : exchangeUser.Name;

            return syncConfigured &&
                   olItem.Sensitivity == OlSensitivity.olNormal &&
                   /* If there is a valid crmId it's arrived via CRM and is therefore safe to save to CRM;
                    * if the current user is the organiser, AND there's no valid CRM id, then it's a new one
                    * that the current user made, and we should save it to CRM. */
                   (CrmId.IsInvalid(olItem.GetCrmId()) || currentUserName == organiser) &&
                   /* Microsoft Conferencing Add-in creates temporary items with names which start 
                    * 'PLEASE IGNORE' - we should not sync these. */
                   !olItem.Subject.StartsWith(MSConfTmpSubjectPrefix);
        }

        /// <summary>
        ///     Synchronise items in the specified folder with the specified SuiteCRM module.
        /// </summary>
        /// <remarks>
        ///     TODO: candidate for refactoring upwards, in concert with ContactSyncing.SyncFolder.
        /// </remarks>
        /// <param name="folder">The folder.</param>
        /// <param name="crmModule">The module.</param>
        protected override void SyncFolder(MAPIFolder folder, string crmModule)
        {
            Log.Debug($"{GetType().Name}.SyncFolder: '{crmModule}'");

            try
            {
                /* this.ItemsSyncState already contains items to be synced. */
                var untouched =
                    new HashSet<SyncState<AppointmentItem>>(SyncStateManager.Instance
                        .GetSynchronisedItems<SyncStateType>());
                var records = MergeRecordsFromCrm(folder, crmModule, untouched);

                AddOrUpdateItemsFromCrmToOutlook(records, folder, untouched, crmModule);

                var invited = RestAPIWrapper.GetRelationships("Users",
                    RestAPIWrapper.GetUserId(), crmModule.ToLower(),
                    RestAPIWrapper.GetSugarFields(crmModule));
                if (invited != null)
                    AddOrUpdateItemsFromCrmToOutlook(invited, folder, untouched, crmModule);

                try
                {
                    ResolveUnmatchedItems(untouched);
                }
                catch (Exception ex)
                {
                    ErrorHandler.Handle($"Failed while synchronising {DefaultCrmModule}", ex);
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle($"Failed while synchronising {DefaultCrmModule}", ex);
            }
        }

        /// <summary>
        ///     Update an existing Outlook item with values taken from a corresponding CRM item. Note that
        ///     this just overwrites all values in the Outlook item.
        /// </summary>
        /// <param name="crmType">The CRM type of the item from which values are to be taken.</param>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <param name="dateStart">The state date/time of the item, adjusted for timezone.</param>
        /// <param name="syncState">The outlook item assumed to correspond with the CRM item.</param>
        /// <returns>An appropriate sync state.</returns>
        private SyncState<AppointmentItem> UpdateExistingOutlookItemFromCrm(
            string crmType,
            EntryValue crmItem,
            DateTime dateStart,
            SyncState<AppointmentItem> syncState)
        {
            LogItemAction(syncState.OutlookItem, "AppointmentSyncing.UpdateExistingOutlookItemFromCrm");

            if (!syncState.IsDeletedInOutlook)
            {
                var olItem = syncState.OutlookItem;
                var olPropertyModifiedDate = olItem!= null && olItem.IsValid() ? olItem.UserProperties[SyncStateManager.ModifiedDatePropertyName] : null;

                if (olPropertyModifiedDate == null || olPropertyModifiedDate.Value !=
                    crmItem.GetValueAsString("date_modified"))
                    try
                    {
                        olItem.Subject = crmItem.GetValueAsString("name");
                        olItem.Body = crmItem.GetValueAsString("description");
                        if (!string.IsNullOrWhiteSpace(crmItem.GetValueAsString("date_start")))
                            UpdateOutlookDetails(crmType, crmItem, dateStart, olItem);

                        EnsureSynchronisationPropertiesForOutlookItem(olItem, crmItem, crmType);
                        LogItemAction(syncState.OutlookItem,
                            "AppointmentSyncing.UpdateExistingOutlookItemFromCrm, item saved");
                    }
                    catch (Exception any)
                    {
                        Globals.ThisAddIn.Log.Warn($"Unexpected error in UpdateExistingOutlookItemFromCrm", any);
                    }
                    finally
                    {
                        SaveItem(olItem);
                    }
                syncState.OModifiedDate =
                    DateTime.ParseExact(crmItem.GetValueAsString("date_modified"), "yyyy-MM-dd HH:mm:ss", null);
            }

            return syncState;
        }

        /// <summary>
        ///     Update this Outlook appointment's start and duration from this CRM object.
        /// </summary>
        /// <param name="crmType">The CRM type of the item from which values are to be taken.</param>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <param name="date_start">The state date/time of the item, adjusted for timezone.</param>
        /// <param name="olItem">The outlook item assumed to correspond with the CRM item.</param>
        protected virtual void UpdateOutlookDetails(string crmType, EntryValue crmItem, DateTime date_start,
            AppointmentItem olItem)
        {
            try
            {
                olItem.Start = date_start;
                var minutesString = crmItem.GetValueAsString("duration_minutes");
                var hoursString = crmItem.GetValueAsString("duration_hours");

                var minutes = string.IsNullOrWhiteSpace(minutesString) ? 0 : int.Parse(minutesString);
                var hours = string.IsNullOrWhiteSpace(hoursString) ? 0 : int.Parse(hoursString);

                olItem.Duration = minutes + hours * 60;
            }

            finally
            {
                SaveItem(olItem);
            }
        }


        internal override string GetOutlookEntryId(AppointmentItem olItem)
        {
            return olItem.EntryID;
        }

        protected override CrmId GetCrmEntryId(AppointmentItem olItem)
        {
            return olItem.GetCrmId();
        }

        /// <summary>
        ///     Return the sensitivity of this outlook item.
        /// </summary>
        /// <remarks>
        ///     Outlook item classes do not inherit from a common base class, so generic client code cannot refer to
        ///     'OutlookItem.Sensitivity'.
        /// </remarks>
        /// <param name="item">The outlook item whose sensitivity is required.</param>
        /// <returns>the sensitivity of the item.</returns>
        internal override OlSensitivity GetSensitivity(AppointmentItem item)
        {
            return item.Sensitivity;
        }


        /// <summary>
        ///     Add an address resolution composed from this module name and record to the cache.
        /// </summary>
        /// <param name="moduleName">The name of the module in which the record was found</param>
        /// <param name="record">The record.</param>
        protected void CacheAddressResolutionData(string moduleName, LinkRecord record)
        {
            CacheAddressResolutionData(
                new AddressResolutionData(moduleName,
                    record.data.AsDictionary()));
        }

        /// <summary>
        ///     Add this resolution to the cache.
        /// </summary>
        /// <param name="resolution">The resolution to add.</param>
        protected void CacheAddressResolutionData(AddressResolutionData resolution)
        {
            List<AddressResolutionData> resolutions;

            if (meetingRecipientsCache.ContainsKey(resolution.EmailAddress))
            {
                resolutions = meetingRecipientsCache[resolution.EmailAddress];
            }
            else
            {
                resolutions = new List<AddressResolutionData>();
                meetingRecipientsCache[resolution.EmailAddress] = resolutions;
            }

            if (!resolutions.Any(x => x.ModuleId == resolution.ModuleId && x.ModuleName == resolution.ModuleName))
                resolutions.Add(resolution);

            Log.Debug(
                $"Successfully cached recipient {resolution.EmailAddress} => {resolution.ModuleName}, {resolution.ModuleId}.");
        }


        protected void CacheAddressResolutionData(EntryValue crmItem)
        {
            foreach (var list in crmItem.relationships.link_list)
            foreach (var record in list.records)
            {
                var data = record.data.AsDictionary();
                try
                {
                    CacheAddressResolutionData(list.name, record);
                }
                catch (Exception e) when (e is TypeInitializationException || e is InvalidCrmIdException)
                {
                    ErrorHandler.Handle("Probable invalid CRM ID", e);
                }
                catch (KeyNotFoundException kex)
                {
                    ErrorHandler.Handle(
                        $"Email address '{record.data.GetValueAsString(AddressResolutionData.EmailAddressFieldName)}' not recognised while caching meeting recipients.",
                        kex);
                }
            }
        }


        /// <summary>
        ///     Used for caching data for resolving email addresses to CRM records.
        /// </summary>
        protected class AddressResolutionData
        {
            /// <summary>
            ///     Expected name in the input map of the email address field.
            /// </summary>
            public const string EmailAddressFieldName = "email1";

            /// <summary>
            ///     Expected name in the input map of the field containing the id in
            ///     the specified module.
            /// </summary>
            public const string ModuleIdFieldName = "id";

            /// <summary>
            ///     Expected name in the input map of the field containing an associated id in
            ///     the `Accounts` module, if any.
            /// </summary>
            public const string AccountIdFieldName = "account_id";

            /// <summary>
            ///     The id within the `Accounts` module of a related record, if any.
            /// </summary>
            private readonly object accountId;

            /// <summary>
            ///     The email address resolved by this data.
            /// </summary>
            public readonly string EmailAddress;

            /// <summary>
            ///     The id within that module of the record to which it resolves.
            /// </summary>
            public readonly CrmId ModuleId;

            /// <summary>
            ///     The name of the CRM module to which it resolves.
            /// </summary>
            public readonly string ModuleName;

            public AddressResolutionData(string moduleName, CrmId moduleId, string emailAddress)
            {
                ModuleName = moduleName;
                ModuleId = moduleId;
                EmailAddress = emailAddress;
            }

            public AddressResolutionData(string moduleName, Dictionary<string, object> data)
            {
                ModuleName = moduleName;
                ModuleId = CrmId.Get(data[ModuleIdFieldName]);
                EmailAddress = data[EmailAddressFieldName]?.ToString();
                try
                {
                    accountId = data[AccountIdFieldName]?.ToString();
                }
                catch (KeyNotFoundException)
                {
                    // and ignore it; that key often won't be there.
                }
            }
        }
    }
}
