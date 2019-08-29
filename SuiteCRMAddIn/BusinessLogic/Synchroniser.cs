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
using System.Globalization;
using System.Linq;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using SuiteCRMAddIn.Daemon;
using SuiteCRMAddIn.Exceptions;
using SuiteCRMAddIn.Properties;
using SuiteCRMClient;
using SuiteCRMClient.Logging;
using SuiteCRMClient.RESTObjects;
using Exception = System.Exception;
using System.Runtime.InteropServices;

#endregion

namespace SuiteCRMAddIn.BusinessLogic
{
    public abstract class Synchroniser : RepeatingProcess
    {
        public Synchroniser(string name, SyncContext context) : base(name, context.Log)
        {
            Context = context;
        }

        protected Application Application => Context.Application;

        /// <summary>
        ///     The name of the default CRM module (record type) that this synchroniser synchronises.
        /// </summary>
        public abstract string DefaultCrmModule { get; }

        /// <summary>
        ///     The direction(s) in which I sync
        /// </summary>
        public abstract SyncDirection.Direction Direction { get; }

        /// <summary>
        ///     The synchronisation context in which I operate.
        /// </summary>
        protected SyncContext Context { get; }

        /// <summary>
        ///     Get the Outlook folder in which my items are stored.
        /// </summary>
        /// <returns>the Outlook folder in which my items are stored</returns>
        public abstract MAPIFolder GetDefaultFolder();


        /// <summary>
        ///     Specialisation: get my Outlook items.
        /// </summary>
        public override void PerformStartup()
        {
            LinkOutlookItems(GetDefaultFolder());
        }

        /// <summary>
        /// Links the outlook items in the specified folder into the items whose synchronisation states I track.
        /// </summary>
        /// <param name="folder">the Outlook folder from which items should be taken.</param>
        protected abstract void LinkOutlookItems(MAPIFolder folder);


        /// <summary>
        /// Synchronise items in the specified folder with the specified SuiteCRM module.
        /// </summary>
        protected abstract void SyncFolder(MAPIFolder folder, string crmModule);
    }

    /// <summary>
    ///     Synchronise items of the class for which I am responsible.
    /// </summary>
    /// <remarks>
    ///     It's arguable that specialisations of this class ought to be singletons, but currently they are not.
    /// </remarks>
    /// <typeparam name="OutlookItemType">The class of item that I am responsible for synchronising.</typeparam>
    /// <typeparam name="SyncStateType">An appropriate sync state type for my <see cref="OutlookItemType" /></typeparam>
    public abstract class Synchroniser<OutlookItemType, SyncStateType> : Synchroniser, IDisposable
        where OutlookItemType : class
        where SyncStateType : SyncState<OutlookItemType>
    {
        /// <summary>
        ///     A cache for CRM premissions to prevent continually asking for them.
        /// </summary>
        protected readonly CRMPermissionsCache<OutlookItemType, SyncStateType> permissionsCache;

        // Keep a reference to the COM object on which we have event handlers, otherwise
        // when the reference is garbage-collected, the event-handlers are removed!
        private Items _itemsCollection;

        /// <summary>
        ///     A lock on the creation of new objects in Outlook.
        /// </summary>
        protected object creationLock = new object();

        /// <summary>
        ///     It appears that CRM sends us back strings HTML escaped.
        /// </summary>
        protected JsonSerializerSettings DeserialiseSettings = new JsonSerializerSettings
        {
            StringEscapeHandling = StringEscapeHandling.EscapeHtml
        };

        /// <summary>
        ///     A lock to prevent enqueueing the same new object twice in different
        ///     threads (unlikely, since it should always be in the VSTA_main thread,
        ///     but let's be paranoid).
        /// </summary>
        protected object enqueueingLock = new object();

        /// <summary>
        ///     The prefix for the fetch query, used in FetchRecordsFromCrm, q.v.
        /// </summary>
        protected string fetchQueryPrefix;

        private string folderName;

        /// <summary>
        ///     Construct a new instance of a synchroniser with this thread name and context.
        /// </summary>
        /// <param name="threadName">The name of the thread I shall create.</param>
        /// <param name="context">The context in which I shall work.</param>
        public Synchroniser(string threadName, SyncContext context) : base(threadName, context)
        {
            InstallEventHandlers();
            AddSuiteCrmOutlookCategory();
            permissionsCache = new CRMPermissionsCache<OutlookItemType, SyncStateType>(this, context.Log);
        }

        public void Dispose()
        {
            // RemoveEventHandlers();
        }

        /// <summary>
        ///     Am I enabled? I.e., am I able either to import or to export?
        /// </summary>
        /// <returns>True if this synchroniser is enabled.</returns>
        public bool IsEnabled()
        {
            return Globals.ThisAddIn.HasCrmUserSession &&
                   (SyncDirection.AllowInbound(Direction) ||
                    SyncDirection.AllowOutbound(Direction));
        }

        /// <summary>
        ///     Get a date stamp for midnight five days ago (why?).
        /// </summary>
        /// <returns>A date stamp for midnight five days ago.</returns>
        public DateTime GetStartDate()
        {
            var dtRet = DateTime.Now.AddDays(-5);
            return new DateTime(dtRet.Year, dtRet.Month, dtRet.Day, 0, 0, 0);
        }

        /// <summary>
        /// Return a SQL fragment based on my start date
        /// </summary>
        /// <returns></returns>
        public string GetStartDateString()
        {
            return $" AND [Start] >='{GetStartDate().ToString("MM/dd/yyyy HH:mm")}'";
        }

        /// <summary>
        /// Prepare me for shutdown by removing my event handlers.
        /// </summary>
        /// <returns>Always zero.</returns>
        public override int PrepareShutdown()
        {
            RemoveEventHandlers();
            return 0;
        }

        /// <summary>
        ///     Run a single iteration of the synchronisation process for the items for which I am responsible.
        /// </summary>
        public virtual void SynchroniseAll()
        {
            Log.Debug($"{GetType().Name} SynchroniseAll starting");

            if (permissionsCache.HasExportAccess())
            {
                var folder = GetDefaultFolder();

                SyncFolder(folder, DefaultCrmModule);
            }
            else
            {
                Log.Debug(
                    $"{GetType().Name}.SynchroniseAll not synchronising {DefaultCrmModule} because export access is denied");
            }

            Log.Debug($"{GetType().Name} SynchroniseAll completed");
        }

        /// <summary>
        ///     Add the item implied by this SyncState, which may not exist in CRM, to CRM.
        /// </summary>
        /// <param name="syncState">The sync state.</param>
        /// <returns>The id of the entry added or updated.</returns>
        internal virtual CrmId AddOrUpdateItemFromOutlookToCrm(SyncState<OutlookItemType> syncState)
        {
            var result = CrmId.Empty;

            lock (enqueueingLock)
            {
                if (!syncState.VerifyItem())
                {
                    // TODO: this puts us into a death spiral, if item is a meeting. 
                    HandleItemMissingFromOutlook(syncState);
                }
                else if (ShouldAddOrUpdateItemFromOutlookToCrm(syncState.OutlookItem))
                {
                    var olItem = syncState.OutlookItem;

                    try
                    {
                        LogItemAction(olItem, "Synchroniser.AddOrUpdateItemFromOutlookToCrm, Despatching");

                        try
                        {
                            syncState.SetTransmitted();

                            result = ConstructAndDespatchCrmItem(syncState);
                            if (CrmId.IsValid(result))
                            {
                                var utcNow = DateTime.UtcNow;
                                EnsureSynchronisationPropertiesForOutlookItem(olItem, utcNow.ToString(),
                                    DefaultCrmModule, result);

                                syncState.SetSynced(result);
                            }
                            else
                            {
                                Log.Warn(
                                    "AppointmentSyncing.AddItemFromOutlookToCrm: Invalid CRM Id returned; item may not be stored.");
                                syncState.SetPending(true);
                            }
                        }
                        catch (BadStateTransition bst)
                        {
                            /* almost certainly the item had been taken out of the queue because
                             * it had changed, and that's OK */
                            if (bst.From != TransmissionState.Pending)
                                Log.Warn($"Unexpected error while transmitting {DefaultCrmModule}.", bst);
                        }
                        catch (Exception any)
                        {
                            Log.Warn($"Unexpected error while transmitting {DefaultCrmModule}.", any);
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorHandler.Handle("Failed while trying to add or update an item from Outlook to CRM", ex);
                        syncState.SetPending(true);
                    }
                    finally
                    {
                        SaveItem(olItem);
                    }
                }
            }

            return result;
        }

        /// <summary>
        ///     Get the entry id of this Outlook item.
        /// </summary>
        /// <remarks>
        ///     Outlook item classes do not inherit from a common base class, so generic client code cannot refer to
        ///     'OutlookItem.EntryId'; hance this rather clumsy mechanism.
        /// </remarks>
        /// <param name="olItem">The Outlook item from which the entry id should be taken.</param>
        /// <returns>the entry id of this Outlook item.</returns>
        internal abstract string GetOutlookEntryId(OutlookItemType olItem);

        /// <summary>
        ///     Return the sensitivity of this outlook item.
        /// </summary>
        /// <remarks>
        ///     Outlook item classes do not inherit from a common base class, so generic client code cannot refer to
        ///     'OutlookItem.Sensitivity'; hance this rather clumsy mechanism.
        /// </remarks>
        /// <param name="olItem">The outlook item whose sensitivity is required.</param>
        /// <returns>the sensitivity of the item.</returns>
        internal abstract OlSensitivity GetSensitivity(OutlookItemType olItem);


        internal IEnumerable<WithRemovableSynchronisationProperties> GetSynchronisedItems()
        {
            return IsEnabled()
                ? SyncStateManager.Instance.GetSynchronisedItems<SyncStateType>()
                : new List<SyncStateType>();
        }

        /// <summary>
        ///     Deal with an item which used to exist in Outlook but which no longer does.
        ///     The default behaviour is to remove it from CRM.
        /// </summary>
        /// <param name="syncState">The dangling syncState of the missing item.</param>
        internal virtual void HandleItemMissingFromOutlook(SyncState<OutlookItemType> syncState)
        {
            RemoveFromCrm(syncState);
            RemoveItemSyncState(syncState);
        }

        /// <summary>
        ///     Log a message regarding this Outlook item, with detail of the item.
        /// </summary>
        /// <param name="olItem">The outlook item.</param>
        /// <param name="message">The message to be logged.</param>
        internal abstract void LogItemAction(OutlookItemType olItem, string message);

        /// <summary>
        ///     If I am currently configured to do so, synchronise the items for which I am
        ///     responsible once.
        /// </summary>
        internal override void PerformIteration()
        {
            if (IsEnabled())
            {
                if (SyncDirection.AllowInbound(Direction))
                    SynchroniseAll();
                else
                    Log.Debug($"{GetType().Name}.SynchroniseAll not running because not enabled");

                OtherIterationActions();
            }
            else
            {
                Log.Debug($"{GetType().Name}.SynchroniseAll not running because no session");
            }
        }

        /// <summary>
        ///     Update a single item in the specified Outlook folder with changes from CRM. If the item
        ///     does not exist, create it.
        /// </summary>
        /// <param name="folder">The folder to synchronise into.</param>
        /// <param name="crmType">The CRM type of the candidate item.</param>
        /// <param name="crmItem">The candidate item from CRM.</param>
        /// <returns>The synchronisation state of the item updated (if it was updated).</returns>
        protected abstract SyncState<OutlookItemType> AddOrUpdateItemFromCrmToOutlook(MAPIFolder folder, string crmType,
            EntryValue crmItem);

        /// <summary>
        ///     Update these items, which may or may not already exist in Outlook.
        /// </summary>
        /// <remarks>
        ///     TODO: It would be much better if, rather than taking `untouched` as a mutable argument,
        ///     this method returned a list of items which weren't identified.
        /// </remarks>
        /// <param name="crmItems">The items to be synchronised.</param>
        /// <param name="folder">The outlook folder to synchronise into.</param>
        /// <param name="untouched">
        ///     A list of sync states of existing items which have
        ///     not yet been synchronised; this list is modified (destructuvely changed)
        ///     by the action of this method.
        /// </param>
        /// <param name="crmType">The CRM record type ('module') to be fetched.</param>
        protected virtual void AddOrUpdateItemsFromCrmToOutlook(
            IList<EntryValue> crmItems,
            MAPIFolder folder,
            HashSet<SyncState<OutlookItemType>> untouched,
            string crmType)
        {
            foreach (var crmItem in crmItems)
                try
                {
                    if (ShouldAddOrUpdateItemFromCrmToOutlook(folder, crmType, crmItem))
                    {
                        var state = AddOrUpdateItemFromCrmToOutlook(folder, crmType, crmItem);
                        if (state != null)
                        {
                            untouched.Remove(state);
                            /* Because Outlook offers us back items in another thread as we modify them
                             * they may already be queued for output before we get here. But they should
                             * ideally not be sent, so we forcibly mark them as synced overriding the
                             * normal flow of the state transition engine. */
                            state.SetSynced(true);
                            LogItemAction(state.OutlookItem,
                                "Synchroniser.AddOrUpdateItemsFromCrmToOutlook, item removed from untouched");
                        }
                    }
                    else
                    {
                        /* even if we shouldn't update it, we should remove it from untouched. */
                        untouched.Remove(SyncStateManager.Instance.GetExistingSyncState(crmItem) as SyncStateType);
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandler.Handle($"Failed while trying to add or update {DefaultCrmModule} from Outlook to CRM",
                        ex);
                }
        }

        /// <summary>
        ///     Construct a JSON packet representing this Outlook item, and despatch it to CRM.
        /// </summary>
        /// <remarks>
        ///     You'd think that with normal object oriented programming you could just implement this
        ///     method here, but because Outlook items are not really objects and don't have a common
        ///     superclass you can't. So it has to be implemented in subclasses.
        /// </remarks>
        /// <param name="syncState">The synchronisation state.</param>
        /// <returns>The CRM id of the object created or modified.</returns>
        protected abstract CrmId ConstructAndDespatchCrmItem(SyncState<OutlookItemType> syncState);

        /// <summary>
        ///     Every Outlook item which is to be synchronised must have a property SOModifiedDate,
        ///     a property SType, and a property SEntryId, referencing respectively the last time it
        ///     was modified, the type of CRM item it is to be synchronised with, and the id of the
        ///     CRM item it is to be synchronised with.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="modifiedDate">The value for the SOModifiedDate property.</param>
        /// <param name="type">The value for the SType property (CRM module name).</param>
        /// <param name="entryId">The value for the SEntryId property (CRM item id).</param>
        protected void EnsureSynchronisationPropertiesForOutlookItem(OutlookItemType olItem, string modifiedDate,
            string type, CrmId entryId)
        {
            try
            {
                EnsureSynchronisationPropertyForOutlookItem(olItem, SyncStateManager.ModifiedDatePropertyName,
                    modifiedDate);
                EnsureSynchronisationPropertyForOutlookItem(olItem, SyncStateManager.TypePropertyName, type);

                if (CrmId.IsValid(entryId))
                {
                    EnsureSynchronisationPropertyForOutlookItem(olItem, SyncStateManager.CrmIdPropertyName, entryId);
                    SyncStateManager.Instance.SetByCrmId(entryId,
                        SyncStateManager.Instance.GetOrCreateSyncState(olItem));
                }

                EnsureSynchronisationPropertyForOutlookItem(olItem, SyncStateManager.ModifiedDatePropertyName,
                    modifiedDate);
                EnsureSynchronisationPropertyForOutlookItem(olItem, SyncStateManager.TypePropertyName, type);
            }
            catch (Exception any)
            {
                Log.Warn($"Unexpected error in EnsureSynchronisationPropertiesForOutlookItem", any);
            }
            finally
            {
                SaveItem(olItem);
            }
        }

        /// <summary>
        ///     Set up synchronisation properties for this outlook item from this CRM item, assuming my default CRM module.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmItem">The CRM item.</param>
        protected virtual void EnsureSynchronisationPropertiesForOutlookItem(OutlookItemType olItem, EntryValue crmItem)
        {
            EnsureSynchronisationPropertiesForOutlookItem(
                olItem,
                crmItem,
                DefaultCrmModule);
        }

        /// <summary>
        ///     Set up synchronisation properties for this outlook item from this CRM item, assuming my default CRM module.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmItem">The CRM item.</param>
        /// <param name="type">The value for the SType property (CRM module name).</param>
        protected virtual void EnsureSynchronisationPropertiesForOutlookItem(OutlookItemType olItem, EntryValue crmItem,
            string type)
        {
            EnsureSynchronisationPropertiesForOutlookItem(
                olItem,
                crmItem.GetValueAsString("date_modified"),
                type,
                CrmId.Get(crmItem.id));
        }

        /// <summary>
        ///     Every Outlook item which is to be synchronised must have a property SOModifiedDate,
        ///     a property SType, and a property SEntryId, referencing respectively the last time it
        ///     was modified, the type of CRM item it is to be synchronised with, and the id of the
        ///     CRM item it is to be synchronised with.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="modifiedDate">The value for the SOModifiedDate property.</param>
        /// <param name="type">The value for the SType property.</param>
        /// <param name="entryId">The value for the SEntryId property.</param>
        protected void EnsureSynchronisationPropertiesForOutlookItem(OutlookItemType olItem, DateTime modifiedDate,
            string type, CrmId entryId)
        {
            EnsureSynchronisationPropertiesForOutlookItem(olItem, modifiedDate.ToString("yyyy-MM-dd HH:mm:ss"), type,
                entryId);
        }

        /// <summary>
        ///     Ensure that this Outlook item has a property of this name with this value.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        protected abstract void EnsureSynchronisationPropertyForOutlookItem(OutlookItemType olItem, string name,
            string value);

        protected void EnsureSynchronisationPropertyForOutlookItem(OutlookItemType olItem, string name, CrmId value)
        {
            EnsureSynchronisationPropertyForOutlookItem(olItem, name, value.ToString());
        }

        /// <summary>
        ///     Find any existing Outlook items which appear to be identical to this CRM item.
        /// </summary>
        /// <param name="crmItem">The CRM item to match.</param>
        /// <returns>A list of matching Outlook items.</returns>
        protected List<SyncState<OutlookItemType>> FindMatches(EntryValue crmItem)
        {
            List<SyncState<OutlookItemType>> result;

            try
            {
                result = SyncStateManager.Instance.GetSynchronisedItems<SyncState<OutlookItemType>>()
                    .Where(a => a.VerifyItem() && IsMatch(a.OutlookItem, crmItem))
                    .ToList();
            }
            catch (Exception any)
            {
                ErrorHandler.Handle($"Failure while checking for items matching id '{crmItem.id}' {Environment.NewLine}", any);
                result = new List<SyncState<OutlookItemType>>();
            }

            return result;
        }

        /// <summary>
        ///     Get the CRM entry id of this item, if it has one and is known.
        /// </summary>
        /// <param name="olItem">The item whose id is saught.</param>
        /// <returns>The id, or null if it is not known.</returns>
        protected abstract CrmId GetCrmEntryId(OutlookItemType olItem);

        /// <summary>
        ///     Fetch the page of entries from this module starting at this offset.
        /// </summary>
        /// <param name="offset">The offset into the resultset at which the page begins.</param>
        /// <returns>A set of entries.</returns>
        protected virtual EntryList GetEntriesPage(int offset)
        {
            return RestAPIWrapper.GetEntryList(DefaultCrmModule,
                string.Format(fetchQueryPrefix, RestAPIWrapper.GetUserId()),
                Settings.Default.SyncMaxRecords, "date_start DESC", offset, false,
                RestAPIWrapper.GetSugarFields(DefaultCrmModule));
        }

        /// <summary>
        ///     Check whether this synchroniser is allowed import access for its default CRM module.
        /// </summary>
        /// <returns>true if this synchroniser is allowed import access for its default CRM module.</returns>
        protected bool HasImportAccess()
        {
            return permissionsCache.HasImportAccess(DefaultCrmModule);
        }

        protected virtual void InstallEventHandlers()
        {
            if (_itemsCollection == null)
            {
                var folder = GetDefaultFolder();
                _itemsCollection = folder.Items;
                folderName = folder.Name;
                Log.Debug("Adding event handlers for folder " + folderName);
                _itemsCollection.ItemAdd += Items_ItemAdd;
                _itemsCollection.ItemChange += Items_ItemChange;
                _itemsCollection.ItemRemove += Items_ItemRemove;
            }
        }

        /// <summary>
        ///     Return true if this Outlook item appears to represent the same item as this CRM item.
        /// </summary>
        /// <remarks>
        ///     Intended to help block howlaround.
        /// </remarks>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmItem">The CRM item.</param>
        /// <returns>true if this Outlook item appears to represent the same item as this CRM item.</returns>
        protected abstract bool IsMatch(OutlookItemType olItem, EntryValue crmItem);

        protected void Items_ItemAdd(object olItem)
        {
            Log.Warn($"Outlook {folderName} ItemAdd");
            try
            {
                OutlookItemAdded(olItem as OutlookItemType);
            }
            catch (Exception problem)
            {
                ErrorHandler.Handle($"Failed to handle an item added to {folderName}", problem);
            }
        }

        protected void Items_ItemChange(object olItem)
        {
            Log.Debug($"Outlook {folderName} ItemChange");
            try
            {
                OutlookItemChanged(olItem as OutlookItemType);
            }
            catch (Exception ex) when (ex is InvalidComObjectException || ex is COMException)
            {
                // invalid item passed into Items_ItemChange, which is odd and 
                // worrying but not much we can do.
                Log.Debug("invalid item passed into Items_ItemChange");
            }
            catch (Exception problem)
            {
                ErrorHandler.Handle($"Failed to handle an item modified in {folderName}", problem);
            }
        }

        protected void Items_ItemRemove()
        {
            Log.Debug($"Outlook {folderName} ItemRemove: entry");
            try
            {
                RemoveDeletedItems();
            }
            catch (Exception problem)
            {
                ErrorHandler.Handle($"Failed to handle item(s) removed from {folderName}", problem);
            }
            Log.Debug($"Outlook {folderName} ItemRemove: exit");
        }

        /// <summary>
        ///     Fetch records in pages from CRM, and merge them into Outlook.
        /// </summary>
        /// <param name="folder">The folder to be synchronised.</param>
        /// <param name="crmModule">The name of the CRM module to synchronise with.</param>
        /// <param name="untouched">A list of all known Outlook items, from which those modified by this method are removed.</param>
        protected virtual IList<EntryValue> MergeRecordsFromCrm(MAPIFolder folder, string crmModule,
            HashSet<SyncState<OutlookItemType>> untouched)
        {
            var thisOffset = 0; // offset of current page of entries
            var nextOffset = 0; // offset of the next page of entries, if any.
            var result = new List<EntryValue>();

            /* get candidates for syncrhonisation from SuiteCRM one page at a time */
            do
            {
                /* update the offset to the offset of the next page */
                thisOffset = nextOffset;

                var entriesPage = GetEntriesPage(thisOffset);
                var entryList = entriesPage.entry_list;

                /* get the offset of the next page */
                nextOffset = entriesPage.next_offset;

                if (entryList != null && entryList.Length > 0)
                {
                    /* it should not be, but it has happened that entry_list has been null */
                    result.AddRange(entryList);
                }
            }
            /* when there are no more entries, we'll get a zero-length entry list and nextOffset
             * will have the same value as thisOffset */
            while (thisOffset != nextOffset);

            return result;
        }

        /// <summary>
        ///     A hook to allow specialisations to do something additional to just syncing in their iterations.
        /// </summary>
        protected virtual void OtherIterationActions()
        {
            // by default do nothing
        }

        /// <summary>
        ///     Entry point from event handler when an item is added in Outlook.
        /// </summary>
        /// <remarks>Should always run in the 'VSTA_main' thread.</remarks>
        /// <remarks>Shouldn't happen here.</remarks>
        /// <param name="olItem">The item that has been added.</param>
        protected virtual void OutlookItemAdded(OutlookItemType olItem)
        {
            if (Globals.ThisAddIn.IsLicensed)
                try
                {
                    OutlookItemAdded(olItem, this);
                }
                catch (Exception any)
                {
                    Log.Warn($"Unexpected error in OutlookItemAdded", any);
                }
                finally
                {
                    if (olItem != null)
                        SaveItem(olItem);
                }
            else
                Log.Warn(
                    $"Synchroniser.OutlookItemAdded: item {GetOutlookEntryId(olItem)} not added because not licensed");
        }

        /// <summary>
        ///     #2246: Nasty workaround for the fact that Outlook 'Appointments' and 'Meetings' are actually the same class.
        /// </summary>
        /// <typeparam name="T">The type of sync state to use.</typeparam>
        /// <param name="olItem">The Outlook item which has been added.</param>
        /// <param name="synchroniser">A synchroniser which can handle the item.</param>
        protected void OutlookItemAdded<T>(OutlookItemType olItem, Synchroniser<OutlookItemType, T> synchroniser)
            where T : SyncState<OutlookItemType>
        {
            LogItemAction(olItem, $"{GetType().Name}.OutlookItemAdded: {GetOutlookEntryId(olItem)}");

            {
                lock (enqueueingLock)
                {
                    if (SyncStateManager.Instance.GetExistingSyncState(olItem) == null)
                    {
                        var state = SyncStateManager.Instance.GetOrCreateSyncState(olItem) as T;
                        if (state != null)
                        {
                            if (olItem != null && ShouldAddOrUpdateItemFromOutlookToCrm(olItem))
                                DaemonWorker.Instance.AddTask(
                                    new TransmitNewAction<OutlookItemType, T>(synchroniser, state));
                        }
                        else
                        {
                            Log.Warn("Should never happen: unexpected sync state type");
                        }
                    }
                    else
                    {
                        Log.Warn(
                            $"{GetType().Name}.OutlookItemAdded: item {GetOutlookEntryId(olItem)} had already been added");
                    }
                }
            }
        }

        /// <summary>
        ///     Entry point from event handler, called when an Outlook item of class AppointmentItem
        ///     is believed to have changed.
        /// </summary>
        /// <param name="olItem">The item which has changed.</param>
        protected virtual void OutlookItemChanged(OutlookItemType olItem)
        {
            if (Globals.ThisAddIn.IsLicensed)
                try
                {
                    OutlookItemChanged(olItem, this);
                }
                catch (BadStateTransition bst)
                {
                    if (bst.From != TransmissionState.Transmitted)
                        throw;
                    /* couldn't set pending -> transmission is in progress */
                }
                finally
                {
                    SaveItem(olItem);
                }
            else
                Log.Warn(
                    $"Synchroniser.OutlookItemAdded: item {GetOutlookEntryId(olItem)} not updated because not licensed");
        }

        /// <summary>
        ///     #2246: Nasty workaround for the fact that Outlook 'Appointments' and 'Meetings' are actually the same class.
        /// </summary>
        /// <typeparam name="T">The type of sync state to use.</typeparam>
        /// <param name="olItem">The Outlook item which has been changed.</param>
        /// <param name="synchroniser">A synchroniser which can handle the item.</param>
        protected void OutlookItemChanged<T>(OutlookItemType olItem, Synchroniser<OutlookItemType, T> synchroniser)
            where T : SyncState<OutlookItemType>
        {
            LogItemAction(olItem, $"{GetType().Name}.OutlookItemChanged: {GetOutlookEntryId(olItem)}");

            SyncState state = SyncStateManager.Instance.GetExistingSyncState(olItem);
            var syncStateForItem = state as T;

            if (syncStateForItem != null)
                try
                {
                    syncStateForItem.SetPending();

                    if (ShouldPerformSyncNow(syncStateForItem))
                        DaemonWorker.Instance.AddTask(
                            new TransmitUpdateAction<OutlookItemType, T>(synchroniser, syncStateForItem));
                    else if (!syncStateForItem.ShouldSyncWithCrm)
                        RemoveFromCrm(syncStateForItem);
                }
                catch (BadStateTransition bst)
                {
                    if (bst.From != TransmissionState.Transmitted)
                        throw;
                }
            else
                OutlookItemAdded(olItem);
        }

        /// <summary>
        ///     Parse a date time object from a user property, assuming the ISO 8601 date-time
        ///     format but ommitting the 'T'. (why? I don't know. TODO: possibly fix).
        /// </summary>
        /// <remarks>
        ///     If the expected format is not recognised, a second scan is attempted without a
        ///     specific format; if this fails, it fails silently and the current time is returned.
        /// </remarks>
        /// <param name="propertyValue">A property value believed to contain a date time string.</param>
        /// <returns>A date time object.</returns>
        protected DateTime ParseDateTimeFromUserProperty(string propertyValue)
        {
            if (propertyValue == null) return default(DateTime);
            var modDateTime = DateTime.UtcNow;
            if (!DateTime.TryParseExact(propertyValue, "yyyy-MM-dd HH:mm:ss", null, DateTimeStyles.None,
                out modDateTime))
                DateTime.TryParse(propertyValue, out modDateTime);
            return modDateTime;
        }

        /// <summary>
        ///     Deal, in CRM, with items deleted in Outlook.
        /// </summary>
        protected void RemoveDeletedItems()
        {
            // Make a copy of the list to avoid mutation error while iterating:
            var syncStatesCopy = SyncStateManager.Instance.GetSynchronisedItems<SyncStateType>();
            foreach (var syncState in syncStatesCopy)
            {
                var shouldDeleteFromCrm = this.IsEnabled() &&
                    SyncDirection.AllowOutbound(this.Direction) && 
                    (syncState.IsDeletedInOutlook || !syncState.VerifyItem() || !syncState.ShouldSyncWithCrm);
                if (shouldDeleteFromCrm) RemoveFromCrm(syncState);
                if (syncState.IsDeletedInOutlook) SyncStateManager.Instance.RemoveSyncState(syncState);
            }
        }

        protected virtual void RemoveEventHandlers()
        {
            if (_itemsCollection != null)
            {
                Log.Debug("Removing event handlers for folder " + folderName);
                _itemsCollection.ItemAdd -= Items_ItemAdd;
                _itemsCollection.ItemChange -= Items_ItemChange;
                _itemsCollection.ItemRemove -= Items_ItemRemove;
                _itemsCollection = null;
            }
        }

        /// <summary>
        ///     Remove the item implied by this sync state from CRM.
        /// </summary>
        /// <param name="state">A sync state wrapping an item which has been deleted or marked private in Outlook.</param>
        protected virtual void RemoveFromCrm(SyncState state)
        {
            if (SyncDirection.AllowOutbound(Direction))
            {
                var crmEntryId = state.CrmEntryId;
                if (state.ExistedInCrm && permissionsCache.HasImportAccess(state.CrmType))
                {
                    var data = new NameValue[2];
                    data[0] = RestAPIWrapper.SetNameValuePair("id", crmEntryId);
                    data[1] = RestAPIWrapper.SetNameValuePair("deleted", "1");
                    RestAPIWrapper.SetEntry(data, state.CrmType);
                }

                state.RemoveCrmLink();
            }
        }

        /// <summary>
        ///     Remove an outlook item and its associated sync state.
        /// </summary>
        /// <param name="syncState">The sync state of the item to remove.</param>
        protected void RemoveItemAndSyncState(SyncState<OutlookItemType> syncState)
        {
            LogItemAction(syncState.OutlookItem, "Synchroniser.RemoveItemAndSyncState, deleting item");
            try
            {
                syncState.DeleteItem();
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle($"Failed while trying to delete a {DefaultCrmModule} item", ex);
            }
            RemoveItemSyncState(syncState);
        }

        /// <summary>
        ///     Remove an item from ItemsSyncState.
        /// </summary>
        /// <param name="item">The sync state of the item to remove.</param>
        protected void RemoveItemSyncState(SyncState<OutlookItemType> item)
        {
            LogItemAction(item.OutlookItem, "Synchroniser.RemoveItemSyncState, removed item from queue");
            SyncStateManager.Instance.RemoveSyncState(item);
        }

        /// <summary>
        ///     Given a list of items which exist in Outlook but are missing from CRM, resolve
        ///     how to handle them.
        /// </summary>
        /// <param name="itemsToResolve">The list of items to resolve.</param>
        protected virtual void ResolveUnmatchedItems(IEnumerable<SyncState<OutlookItemType>> itemsToResolve)
        {
            try
            {
                foreach (var unresolved in itemsToResolve)
                    switch (unresolved.TxState)
                    {
                        case TransmissionState.PendingDeletion:
                            /* If it's to resolve and marked pending deletion, we delete it
                             * (unresolved on two successive iterations): */
                            RemoveItemAndSyncState(unresolved);
                            break;

                        case TransmissionState.Synced:
                            if (unresolved.ExistedInCrm)
                                unresolved.SetPendingDeletion();
                            break;

                        case TransmissionState.Pending:
                        case TransmissionState.PresentAtStartup:
                            if (unresolved.ShouldSyncWithCrm)
                                try
                                {
                                    /* if it's unresolved, pending, and should be synced send it. */
                                    unresolved.SetQueued();
                                    AddOrUpdateItemFromOutlookToCrm(unresolved);
                                }
                                catch (BadStateTransition)
                                {
                                    // ignore.
                                }
                            break;

                        case TransmissionState.Queued:
                            if (unresolved.ShouldSyncWithCrm)
                                try
                                {
                                    /* if it's queued and should be synced send it. */
                                    AddOrUpdateItemFromOutlookToCrm(unresolved);
                                }
                                catch (BadStateTransition bst)
                                {
                                    ErrorHandler.Handle($"Failure while seeking to resolve unmatched items", bst);
                                }
                            break;

                        default:
                            try
                            {
                                unresolved.SetPending();
                            }
                            catch (BadStateTransition bst)
                            {
                                if (bst.From != TransmissionState.Transmitted)
                                    ErrorHandler.Handle($"Failure while seeking to resolve unmatched items", bst);
                            }
                            break;
                    }
            }
            finally
            {
                foreach (SyncState resolved in SyncStateManager.Instance.GetSynchronisedItems<SyncStateType>()
                        .Where(s => s.TxState == TransmissionState.PendingDeletion &&
                                    !itemsToResolve.Contains(s)))
                    /* finally, if there exists an item which had been marked pending deletion, but it has
                     *  been found in CRM (i.e. not in unresolved), mark it as synced */
                    ((SyncState<OutlookItemType>)resolved).SetSynced();
            }
        }

        /// <summary>
        ///     Save this item.
        /// </summary>
        /// <remarks>
        ///     <para>
        ///         Because Outlook items are not proper objects, you cannot call the Save method of
        ///         an Outlook item without knowing its exact class explicitly. So there are what look
        ///         like redundant specialisations of this method; they aren't.
        ///     </para>
        ///     <para>Note that we must be able to save an item whether or not the synchroniser is enabled.</para>
        /// </remarks>
        /// <param name="olItem">The item to save.</param>
        protected abstract void SaveItem(OutlookItemType olItem);

        /// <summary>
        ///     Specialisations should return false if there's a good reason why we should
        ///     NOT sync this item.
        /// </summary>
        /// <param name="folder">The folder to synchronise into.</param>
        /// <param name="crmType">The CRM type of the candidate item.</param>
        /// <param name="crmItem">The candidate item from CRM.</param>
        /// <returns>true</returns>
        protected virtual bool ShouldAddOrUpdateItemFromCrmToOutlook(MAPIFolder folder, string crmType,
            EntryValue crmItem)
        {
            return true;
        }

        /// <summary>
        ///     Perform all the necessary checking before adding or updating an item on CRM.
        /// </summary>
        /// <remarks>
        ///     TODO TODO TODO: This does NOT actually do all the checking. Checking is also
        ///     done in SyncState.ShouldSyncWithCRM, and possibly other places. Fix.
        /// </remarks>
        /// <param name="olItem">
        ///     The item we may seek to add or update, presumed to be of
        ///     my default item type.
        /// </param>
        /// <returns>true if we may attempt to add or update that item.</returns>
        protected virtual bool ShouldAddOrUpdateItemFromOutlookToCrm(OutlookItemType olItem)
        {
            bool result;
            var prefix = "Synchoniser.ShouldAddOrUpdateItemFromOutlookToCrm";

            try
            {
                if (olItem == null)
                {
                    Log.Warn($"{prefix}: attempt to send null to {DefaultCrmModule}?");
                    result = false;
                }
                else
                {
                    if (IsManualOverride(olItem))
                    {
                        Log.Info(
                            $"{prefix}: {DefaultCrmModule} added to CRM because it is set to manual override.");
                        result = true;
                    }
                    else if (SyncDirection.AllowOutbound(Direction))
                    {
                        if (permissionsCache.HasImportAccess(DefaultCrmModule))
                        {
                            if (GetSensitivity(olItem) == OlSensitivity.olNormal)
                            {
                                result = true;
                            }
                            else
                            {
                                Log.Info(
                                    $"{prefix}: {DefaultCrmModule} not added to CRM because its sensitivity is not public.");
                                result = false;
                            }
                        }
                        else
                        {
                            Log.Info(
                                $"{prefix}: {DefaultCrmModule} not added to CRM because import access is not granted.");
                            result = false;
                        }
                    }
                    else
                    {
                        Log.Info(
                            $"{prefix}: {DefaultCrmModule} not added to CRM because synchronisation is not enabled.");
                        result = false;
                    }
                }
            }
            catch (Exception any)
            {
                ErrorHandler.Handle($"Unexpected failure while checking {DefaultCrmModule}.", any);
                result = false;
            }

            return result;
        }

        /// <summary>
        /// Returns true if this `olItem` has a manual override allowing it to be synced while synchronisation
        ///  is disabled.
        /// </summary>
        /// <remarks>
        /// #4754: We need to allow Contacts (but not, at present, other items) to be manually synced even when
        /// synchronisation is otherwise disabled. 
        /// </remarks>
        /// <param name="olItem">The outlook item</param>
        /// <returns>true if this `olItem` has a manual override.</returns>
        protected virtual bool IsManualOverride(OutlookItemType olItem)
        {
            return false;
        }


        /// <summary>
        ///     Should the item represented by this sync state be synchronised now?
        /// </summary>
        /// <param name="syncState">The sync state under consideration.</param>
        /// <returns>True if this synchroniser relates to the current tab and the timing logic is satisfied.</returns>
        protected bool ShouldPerformSyncNow(SyncState<OutlookItemType> syncState)
        {
            return syncState.ShouldPerformSyncNow();
        }

        /// <summary>
        ///     Add the magic 'SuiteCRM' category to the Outlook mapi namespace, if it does not
        ///     already exist.
        /// </summary>
        private void AddSuiteCrmOutlookCategory()
        {
            var oNS = Application.GetNamespace("mapi");
            if (oNS.Categories["SuiteCRM"] == null)
                oNS.Categories.Add("SuiteCRM", OlCategoryColor.olCategoryColorGreen,
                    OlCategoryShortcutKey.olCategoryShortcutKeyNone);
        }
    }
}
