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
    using Daemon;
    using Exceptions;
    using Newtonsoft.Json;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public abstract class Synchroniser : RepeatingProcess
    {
        public Synchroniser(string name, SyncContext context) : base(name, context.Log)
        {
            this.Context = context;
        }

        protected Outlook.Application Application => Context.Application;

        /// <summary>
        /// The name of the default CRM module (record type) that this synchroniser synchronises.
        /// </summary>
        public abstract string DefaultCrmModule
        {
            get;
        }

        /// <summary>
        /// The direction(s) in which I sync
        /// </summary>
        public abstract SyncDirection.Direction Direction { get; }

        /// <summary>
        /// The synchronisation context in which I operate.
        /// </summary>
        protected SyncContext Context { get; private set; }

        /// <summary>
        /// Get the Outlook folder in which my items are stored.
        /// </summary>
        /// <returns></returns>
        public abstract Outlook.MAPIFolder GetDefaultFolder();


        /// <summary>
        /// Specialisation: get my Outlook items.
        /// </summary>
        public override void PerformStartup()
        {
            this.GetOutlookItems(this.GetDefaultFolder());
        }

        protected abstract void GetOutlookItems(Outlook.MAPIFolder folder);

        protected abstract void SyncFolder(Outlook.MAPIFolder folder, string crmModule);
    }

    /// <summary>
    /// Synchronise items of the class for which I am responsible.
    /// </summary>
    /// <remarks>
    /// It's arguable that specialisations of this class ought to be singletons, but currently they are not.
    /// </remarks>
    /// <typeparam name="OutlookItemType">The class of item that I am responsible for synchronising.</typeparam>
    public abstract class Synchroniser<OutlookItemType, SyncStateType> : Synchroniser, IDisposable
        where OutlookItemType : class
        where SyncStateType : SyncState<OutlookItemType>
    {
        /// <summary>
        /// A cache for CRM premissions to prevent continually asking for them.
        /// </summary>
        protected readonly CRMPermissionsCache<OutlookItemType, SyncStateType> permissionsCache;

        /// <summary>
        /// A lock on the creation of new objects in Outlook.
        /// </summary>
        protected object creationLock = new object();

        /// <summary>
        /// It appears that CRM sends us back strings HTML escaped.
        /// </summary>
        protected JsonSerializerSettings deserialiseSettings = new JsonSerializerSettings()
        {
            StringEscapeHandling = StringEscapeHandling.EscapeHtml
        };

        /// <summary>
        /// A lock to prevent enqueueing the same new object twice in different
        /// threads (unlikely, since it should always be in the VSTA_main thread,
        /// but let's be paranoid).
        /// </summary>
        protected object enqueueingLock = new object();

        /// <summary>
        /// The prefix for the fetch query, used in FetchRecordsFromCrm, q.v.
        /// </summary>
        protected string fetchQueryPrefix;

        private string _folderName;

        // Keep a reference to the COM object on which we have event handlers, otherwise
        // when the reference is garbage-collected, the event-handlers are removed!
        private Outlook.Items _itemsCollection = null;

        /// <summary>
        /// Construct a new instance of a synchroniser with this thread name and context.
        /// </summary>
        /// <param name="threadName">The name of the thread I shall create.</param>
        /// <param name="context">The context in which I shall work.</param>
        public Synchroniser(string threadName, SyncContext context) : base( threadName, context)
        {
            this.InstallEventHandlers();
            this.AddSuiteCrmOutlookCategory();
            this.permissionsCache = new CRMPermissionsCache<OutlookItemType, SyncStateType>(this, context.Log);
        }

        public void Dispose()
        {
            // RemoveEventHandlers();
        }

        /// <summary>
        /// Get a date stamp for midnight five days ago (why?).
        /// </summary>
        /// <returns>A date stamp for midnight five days ago.</returns>
        public DateTime GetStartDate()
        {
            DateTime dtRet = DateTime.Now.AddDays(-5);
            return new DateTime(dtRet.Year, dtRet.Month, dtRet.Day, 0, 0, 0);
        }

        public string GetStartDateString()
        {
            return " AND [Start] >='" + GetStartDate().ToString("MM/dd/yyyy HH:mm") + "'";
        }

        public override int PrepareShutdown()
        {
            this.RemoveEventHandlers();
            return 0;
        }

        /// <summary>
        /// Run a single iteration of the synchronisation process for the items for which I am responsible.
        /// </summary>
        public virtual void SynchroniseAll()
        {
            Log.Debug($"{this.GetType().Name} SynchroniseAll starting");

            if (this.permissionsCache.HasExportAccess())
            {
                Outlook.MAPIFolder folder = GetDefaultFolder();

                SyncFolder(folder, this.DefaultCrmModule);
            }
            else
            {
                Log.Debug($"{this.GetType().Name}.SynchroniseAll not synchronising {this.DefaultCrmModule} because export access is denied");
            }

            Log.Debug($"{this.GetType().Name} SynchroniseAll completed");
        }

        /// <summary>
        /// Add the item implied by this SyncState, which may not exist in CRM, to CRM.
        /// </summary>
        /// <param name="syncState">The sync state.</param>
        /// <returns>The id of the entry added or updated.</returns>
        internal virtual string AddOrUpdateItemFromOutlookToCrm(SyncState<OutlookItemType> syncState)
        {
            string result = string.Empty;

            if (this.ShouldAddOrUpdateItemFromOutlookToCrm(syncState.OutlookItem))
            {
                OutlookItemType olItem = syncState.OutlookItem;

                try
                {
                    lock (this.enqueueingLock)
                    {
                        LogItemAction(olItem, "Synchroniser.AddOrUpdateItemFromOutlookToCrm, Despatching");

                        try
                        {
                            syncState.SetTransmitted();

                            result = ConstructAndDespatchCrmItem(olItem);
                            if (!string.IsNullOrEmpty(result))
                            {
                                var utcNow = DateTime.UtcNow;
                                EnsureSynchronisationPropertiesForOutlookItem(olItem, utcNow.ToString(), this.DefaultCrmModule, result);

                                syncState.SetSynced(result);
                            }
                            else
                            {
                                Log.Warn("AppointmentSyncing.AddItemFromOutlookToCrm: Invalid CRM Id returned; item may not be stored.");
                                syncState.SetPending(true);
                            }
                        }
                        catch (BadStateTransition bst)
                        {
                            /* almost certainly the item had been taken out of the queue because
                             * it had changed, and that's OK */
                            if (bst.From != TransmissionState.Pending)
                            {
                                Log.Warn($"Unexpected error while transmitting {this.DefaultCrmModule}.", bst);
                            }
                        }
                        catch (Exception any)
                        {
                            Log.Warn($"Unexpected error while transmitting {this.DefaultCrmModule}.", any);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("Synchroniser.AddOrUpdateItemFromOutlookToCrm", ex);
                    syncState.SetPending(true);
                }
                finally
                {
                    this.SaveItem(olItem);
                }
            }

            return result;
        }

        /// <summary>
        /// Get the entry id of this Outlook item.
        /// </summary>
        /// <param name="olItem">The Outlook item from which the entry id should be taken.</param>
        /// <returns>the entry id of this Outlook item.</returns>
        internal abstract string GetOutlookEntryId(OutlookItemType olItem);

        /// <summary>
        /// Return the sensitivity of this outlook item.
        /// </summary>
        /// <remarks>
        /// Outlook item classes do not inherit from a common base class, so generic client code cannot refer to 'OutlookItem.Sensitivity'.
        /// </remarks>
        /// <param name="olItem">The outlook item whose sensitivity is required.</param>
        /// <returns>the sensitivity of the item.</returns>
        internal abstract Outlook.OlSensitivity GetSensitivity(OutlookItemType olItem);

        internal IEnumerable<WithRemovableSynchronisationProperties> GetSynchronisedItems()
        {
            return SyncStateManager.Instance.GetSynchronisedItems<SyncStateType>(); ;
        }

        /// <summary>
        /// Deal with an item which used to exist in Outlook but which no longer does.
        /// The default behaviour is to remove it from CRM.
        /// </summary>
        /// <param name="syncState">The dangling syncState of the missing item.</param>
        internal virtual void HandleItemMissingFromOutlook(SyncState<OutlookItemType> syncState)
        {
            this.RemoveFromCrm(syncState);
            this.RemoveItemSyncState(syncState);
        }

        /// <summary>
        /// Log a message regarding this Outlook item, with detail of the item.
        /// </summary>
        /// <param name="olItem">The outlook item.</param>
        /// <param name="message">The message to be logged.</param>
        internal abstract void LogItemAction(OutlookItemType olItem, string message);

        /// <summary>
        /// If I am currently configured to do so, synchronise the items for which I am
        /// responsible once.
        /// </summary>
        internal override void PerformIteration()
        {
            if (Globals.ThisAddIn.HasCrmUserSession)
            {
                if (SyncDirection.AllowInbound(this.Direction))
                {
                    this.SynchroniseAll();
                }
                else
                {
                    Log.Debug($"{this.GetType().Name}.SynchroniseAll not running because not enabled");
                }

                this.OtherIterationActions();
            }
            else
            {
                Log.Debug($"{this.GetType().Name}.SynchroniseAll not running because no session");
            }
        }

        /// <summary>
        /// Update a single item in the specified Outlook folder with changes from CRM. If the item
        /// does not exist, create it.
        /// </summary>
        /// <param name="folder">The folder to synchronise into.</param>
        /// <param name="crmType">The CRM type of the candidate item.</param>
        /// <param name="crmItem">The candidate item from CRM.</param>
        /// <returns>The synchronisation state of the item updated (if it was updated).</returns>
        protected abstract SyncState<OutlookItemType> AddOrUpdateItemFromCrmToOutlook(Outlook.MAPIFolder folder, string crmType, EntryValue crmItem);

        /// <summary>
        /// Update these items, which may or may not already exist in Outlook.
        /// </summary>
        /// <remarks>
        /// TODO: It would be much better if, rather than taking `untouched` as a mutable argument,
        /// this method returned a list of items which weren't identified.
        /// </remarks>
        /// <param name="crmItems">The items to be synchronised.</param>
        /// <param name="folder">The outlook folder to synchronise into.</param>
        /// <param name="untouched">A list of sync states of existing items which have
        /// not yet been synchronised; this list is modified (destructuvely changed)
        /// by the action of this method.</param>
        /// <param name="crmType">The CRM record type ('module') to be fetched.</param>
        protected virtual void AddOrUpdateItemsFromCrmToOutlook(
            IList<EntryValue> crmItems,
            Outlook.MAPIFolder folder,
            HashSet<SyncState<OutlookItemType>> untouched,
            string crmType)
        {
            foreach (var crmItem in crmItems)
            {
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
                            LogItemAction(state.OutlookItem, "Synchroniser.AddOrUpdateItemsFromCrmToOutlook, item removed from untouched");
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
                    Log.Error("Synchroniser.AddOrUpdateItemsFromCrmToOutlook", ex);
                }
            }
        }

        /// <summary>
        /// Construct a JSON packet representing this Outlook item, and despatch it to CRM.
        /// </summary>
        /// <remarks>
        /// You'd think that with normal object oriented programming you could just implement this
        /// method here, but because Outlook items are not really objects and don't have a common
        /// superclass you can't. So it has to be implemented in subclasses.
        /// </remarks>
        /// <param name="olItem">The Outlook item.</param>
        /// <returns>The CRM id of the object created or modified.</returns>
        protected abstract string ConstructAndDespatchCrmItem(OutlookItemType olItem);

        /// <summary>
        /// Every Outlook item which is to be synchronised must have a property SOModifiedDate,
        /// a property SType, and a property SEntryId, referencing respectively the last time it
        /// was modified, the type of CRM item it is to be synchronised with, and the id of the
        /// CRM item it is to be synchronised with.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="modifiedDate">The value for the SOModifiedDate property.</param>
        /// <param name="type">The value for the SType property (CRM module name).</param>
        /// <param name="entryId">The value for the SEntryId property (CRM item id).</param>
        protected void EnsureSynchronisationPropertiesForOutlookItem(OutlookItemType olItem, string modifiedDate, string type, string entryId)
        {
            try
            {
                EnsureSynchronisationPropertyForOutlookItem(olItem, SyncStateManager.ModifiedDatePropertyName, modifiedDate);
                EnsureSynchronisationPropertyForOutlookItem(olItem, SyncStateManager.TypePropertyName, type);
                EnsureSynchronisationPropertyForOutlookItem(olItem, SyncStateManager.CrmIdPropertyName, entryId);

                if (!string.IsNullOrEmpty(entryId))
                {
                    SyncStateManager.Instance.SetByCrmId(entryId, SyncStateManager.Instance.GetOrCreateSyncState(olItem));
                }
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
        /// Set up synchronisation properties for this outlook item from this CRM item, assuming my default CRM module.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmItem">The CRM item.</param>
        protected virtual void EnsureSynchronisationPropertiesForOutlookItem(OutlookItemType olItem, EntryValue crmItem)
        {
            this.EnsureSynchronisationPropertiesForOutlookItem(
                olItem,
                crmItem,
                this.DefaultCrmModule);
        }

        /// <summary>
        /// Set up synchronisation properties for this outlook item from this CRM item, assuming my default CRM module.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmItem">The CRM item.</param>
        /// <param name="type">The value for the SType property (CRM module name).</param>
        protected virtual void EnsureSynchronisationPropertiesForOutlookItem(OutlookItemType olItem, EntryValue crmItem, string type)
        {
            this.EnsureSynchronisationPropertiesForOutlookItem(
                olItem,
                crmItem.GetValueAsString("date_modified"),
                type,
                crmItem.GetValueAsString("id"));
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
        protected void EnsureSynchronisationPropertiesForOutlookItem(OutlookItemType olItem, DateTime modifiedDate, string type, string entryId)
        {
            this.EnsureSynchronisationPropertiesForOutlookItem(olItem, modifiedDate.ToString("yyyy-MM-dd HH:mm:ss"), type, entryId);
        }

        /// <summary>
        /// Ensure that this Outlook item has a property of this name with this value.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        protected abstract void EnsureSynchronisationPropertyForOutlookItem(OutlookItemType olItem, string name, string value);

        /// <summary>
        /// Find any existing Outlook items which appear to be identical to this CRM item.
        /// </summary>
        /// <param name="crmItem">The CRM item to match.</param>
        /// <returns>A list of matching Outlook items.</returns>
        protected List<SyncState<OutlookItemType>> FindMatches(EntryValue crmItem)
        {
            List<SyncState<OutlookItemType>> result;

            try
            {
                result = SyncStateManager.Instance.GetSynchronisedItems<SyncState<OutlookItemType>>().Where(a => this.IsMatch(a.OutlookItem, crmItem))
                    .ToList<SyncState<OutlookItemType>>();
            }
            catch (Exception any)
            {
                this.Log.Error("Exception while checking for matches", any);
                result = new List<SyncState<OutlookItemType>>();
            }

            return result;
        }

        /// <summary>
        /// Get the CRM entry id of this item, if it has one and is known.
        /// </summary>
        /// <param name="olItem">The item whose id is saught.</param>
        /// <returns>The id, or null if it is not known.</returns>
        protected abstract string GetCrmEntryId(OutlookItemType olItem);

        /// <summary>
        /// Fetch the page of entries from this module starting at this offset.
        /// </summary>
        /// <param name="offset">The offset into the resultset at which the page begins.</param>
        /// <returns>A set of entries.</returns>
        protected virtual EntryList GetEntriesPage(int offset)
        {
            return RestAPIWrapper.GetEntryList(this.DefaultCrmModule,
                String.Format(fetchQueryPrefix, RestAPIWrapper.GetUserId()),
                Properties.Settings.Default.SyncMaxRecords, "date_start DESC", offset, false,
                RestAPIWrapper.GetSugarFields(this.DefaultCrmModule));
        }

        /// <summary>
        /// Check whether this synchroniser is allowed import access for its default CRM module.
        /// </summary>
        /// <returns>true if this synchroniser is allowed import access for its default CRM module.</returns>
        protected bool HasImportAccess()
        {
            return this.permissionsCache.HasImportAccess(this.DefaultCrmModule);
        }

        protected virtual void InstallEventHandlers()
        {
            if (_itemsCollection == null)
            {
                var folder = GetDefaultFolder();
                _itemsCollection = folder.Items;
                _folderName = folder.Name;
                Log.Debug("Adding event handlers for folder " + _folderName);
                _itemsCollection.ItemAdd += Items_ItemAdd;
                _itemsCollection.ItemChange += Items_ItemChange;
                _itemsCollection.ItemRemove += Items_ItemRemove;
            }
        }

        /// <summary>
        /// Return true if this Outlook item appears to represent the same item as this CRM item.
        /// </summary>
        /// <remarks>
        /// Intended to help block howlaround.
        /// </remarks>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmItem">The CRM item.</param>
        /// <returns>true if this Outlook item appears to represent the same item as this CRM item.</returns>
        protected abstract bool IsMatch(OutlookItemType olItem, EntryValue crmItem);

        protected void Items_ItemAdd(object olItem)
        {
            Log.Warn($"Outlook {_folderName} ItemAdd");
            try
            {
                OutlookItemAdded(olItem as OutlookItemType);
            }
            catch (Exception problem)
            {
                Log.Error($"{_folderName} ItemAdd failed", problem);
            }
        }

        protected void Items_ItemChange(object olItem)
        {
            Log.Debug($"Outlook {_folderName} ItemChange");
            try
            {
                OutlookItemChanged(olItem as OutlookItemType);
            }
            catch (Exception problem)
            {
                Log.Error($"{_folderName} ItemChange failed", problem);
            }
        }

        protected void Items_ItemRemove()
        {
            Log.Debug($"Outlook {_folderName} ItemRemove");
            try
            {
                RemoveDeletedItems();
            }
            catch (Exception problem)
            {
                Log.Error($"{_folderName} ItemRemove failed", problem);
            }
        }

        /// <summary>
        /// Fetch records in pages from CRM, and merge them into Outlook.
        /// </summary>
        /// <param name="folder">The folder to be synchronised.</param>
        /// <param name="crmModule">The name of the CRM module to synchronise with.</param>
        /// <param name="untouched">A list of all known Outlook items, from which those modified by this method are removed.</param>
        protected virtual IList<EntryValue> MergeRecordsFromCrm(Outlook.MAPIFolder folder, string crmModule, HashSet<SyncState<OutlookItemType>> untouched)
        {
            int thisOffset = 0; // offset of current page of entries
            int nextOffset = 0; // offset of the next page of entries, if any.
            List<EntryValue> result = new List<EntryValue>();

            /* get candidates for syncrhonisation from SuiteCRM one page at a time */
            do
            {
                /* update the offset to the offset of the next page */
                thisOffset = nextOffset;

                EntryList entriesPage = GetEntriesPage(thisOffset);

                /* get the offset of the next page */
                nextOffset = entriesPage.next_offset;

                result.AddRange(entriesPage.entry_list);
            }
            /* when there are no more entries, we'll get a zero-length entry list and nextOffset
             * will have the same value as thisOffset */
            while (thisOffset != nextOffset);

            return result;
        }

        /// <summary>
        /// A hook to allow specialisations to do something additional to just syncing in their iterations.
        /// </summary>
        protected virtual void OtherIterationActions()
        {
            // by default do nothing
        }

        /// <summary>
        /// Entry point from event handler when an item is added in Outlook.
        /// </summary>
        /// <remarks>Should always run in the 'VSTA_main' thread.</remarks>
        /// <remarks>Shouldn't happen here.</remarks>
        /// <param name="olItem">The item that has been added.</param>
        protected virtual void OutlookItemAdded(OutlookItemType olItem)
        {
            if (Globals.ThisAddIn.IsLicensed)
            {
                try
                {
                    OutlookItemAdded<SyncStateType>(olItem, this);
                }
                catch (Exception any)
                {
                    Log.Warn($"Unexpected error in OutlookItemAdded", any);
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
        /// #2246: Nasty workaround for the fact that Outlook 'Appointments' and 'Meetings' are actually the same class.
        /// </summary>
        /// <typeparam name="T">The type of sync state to use.</typeparam>
        /// <param name="olItem">The Outlook item which has been added.</param>
        /// <param name="synchroniser">A synchroniser which can handle the item.</param>
        protected void OutlookItemAdded<T>(OutlookItemType olItem, Synchroniser<OutlookItemType, T> synchroniser)
            where T : SyncState<OutlookItemType>
        {
            LogItemAction(olItem, $"{this.GetType().Name}.OutlookItemAdded: {this.GetOutlookEntryId(olItem)}");

            {
                lock (enqueueingLock)
                {
                    if (SyncStateManager.Instance.GetExistingSyncState(olItem) == null)
                    {
                        T state = SyncStateManager.Instance.GetOrCreateSyncState(olItem) as T;
                        if (state != null)
                        {
                            if (olItem != null && this.ShouldAddOrUpdateItemFromOutlookToCrm(olItem))
                            {
                                DaemonWorker.Instance.AddTask(new TransmitNewAction<OutlookItemType, T>(synchroniser, state));
                            }
                        }
                        else
                        {
                            Log.Warn("Should never happen: unexpected sync state type");
                        }
                    }
                    else
                    {
                        Log.Warn($"{this.GetType().Name}.OutlookItemAdded: item {this.GetOutlookEntryId(olItem)} had already been added");
                    }
                }
            }
        }

        /// <summary>
        /// Entry point from event handler, called when an Outlook item of class AppointmentItem
        /// is believed to have changed.
        /// </summary>
        /// <param name="olItem">The item which has changed.</param>
        protected virtual void OutlookItemChanged(OutlookItemType olItem)
        {
            if (Globals.ThisAddIn.IsLicensed)
            {
                try
                {
                    OutlookItemChanged<SyncStateType>(olItem, this);
                }
                catch (BadStateTransition bst)
                {
                    if (bst.From != TransmissionState.Transmitted)
                    {
                        throw;
                    }
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

        /// <summary>
        /// #2246: Nasty workaround for the fact that Outlook 'Appointments' and 'Meetings' are actually the same class.
        /// </summary>
        /// <typeparam name="T">The type of sync state to use.</typeparam>
        /// <param name="olItem">The Outlook item which has been changed.</param>
        /// <param name="synchroniser">A synchroniser which can handle the item.</param>
        protected void OutlookItemChanged<T>(OutlookItemType olItem, Synchroniser<OutlookItemType, T> synchroniser)
            where T : SyncState<OutlookItemType>
        {
            LogItemAction(olItem, $"{this.GetType().Name}.OutlookItemChanged: {this.GetOutlookEntryId(olItem)}");

            SyncState state = SyncStateManager.Instance.GetExistingSyncState(olItem);
            T syncStateForItem = state as T;

            if (syncStateForItem != null)
            {
                try
                {
                    syncStateForItem.SetPending();

                    if (this.ShouldPerformSyncNow(syncStateForItem))
                    {
                        DaemonWorker.Instance.AddTask(new TransmitUpdateAction<OutlookItemType, T>(synchroniser, syncStateForItem));
                    }
                    else if (!syncStateForItem.ShouldSyncWithCrm)
                    {
                        this.RemoveFromCrm(syncStateForItem);
                    }
                }
                catch (BadStateTransition bst)
                {
                    if (bst.From != TransmissionState.Transmitted)
                    {
                        throw;
                    }
                }
            }
            else
            {
                /* we don't have a sync state for this item (presumably formerly private);
                 * that's OK, treat it as new */
                OutlookItemAdded(olItem);
            }
        }

        /// <summary>
        /// Parse a date time object from a user property, assuming the ISO 8601 date-time
        /// format but ommitting the 'T'. (why? I don't know. TODO: possibly fix).
        /// </summary>
        /// <remarks>
        /// If the expected format is not recognised, a second scan is attempted without a
        /// specific format; if this fails, it fails silently and the current time is returned.
        /// </remarks>
        /// <param name="propertyValue">A property value believed to contain a date time string.</param>
        /// <returns>A date time object.</returns>
        protected DateTime ParseDateTimeFromUserProperty(string propertyValue)
        {
            if (propertyValue == null) return default(DateTime);
            var modDateTime = DateTime.UtcNow;
            if (!DateTime.TryParseExact(propertyValue, "yyyy-MM-dd HH:mm:ss", null, DateTimeStyles.None, out modDateTime))
            {
                DateTime.TryParse(propertyValue, out modDateTime);
            }
            return modDateTime;
        }

        /// <summary>
        /// Deal, in CRM, with items deleted in Outlook.
        /// </summary>
        protected void RemoveDeletedItems()
        {
            // Make a copy of the list to avoid mutation error while iterating:
            var syncStatesCopy = SyncStateManager.Instance.GetSynchronisedItems<SyncStateType>();
            foreach (var syncState in syncStatesCopy)
            {
                var shouldDeleteFromCrm = syncState.IsDeletedInOutlook || !syncState.ShouldSyncWithCrm;
                if (shouldDeleteFromCrm) { RemoveFromCrm(syncState); }
                if (syncState.IsDeletedInOutlook) { SyncStateManager.Instance.RemoveSyncState(syncState); }
            }
        }

        protected virtual void RemoveEventHandlers()
        {
            if (_itemsCollection != null)
            {
                Log.Debug("Removing event handlers for folder " + _folderName);
                _itemsCollection.ItemAdd -= Items_ItemAdd;
                _itemsCollection.ItemChange -= Items_ItemChange;
                _itemsCollection.ItemRemove -= Items_ItemRemove;
                _itemsCollection = null;
            }
        }

        /// <summary>
        /// Remove the item implied by this sync state from CRM.
        /// </summary>
        /// <param name="state">A sync state wrapping an item which has been deleted or marked private in Outlook.</param>
        protected virtual void RemoveFromCrm(SyncState state)
        {
            if (SyncDirection.AllowOutbound(Direction))
            {
                var crmEntryId = state.CrmEntryId;
                if (state.ExistedInCrm && this.permissionsCache.HasImportAccess(state.CrmType))
                {
                    NameValue[] data = new NameValue[2];
                    data[0] = RestAPIWrapper.SetNameValuePair("id", crmEntryId);
                    data[1] = RestAPIWrapper.SetNameValuePair("deleted", "1");
                    RestAPIWrapper.SetEntry(data, state.CrmType);
                }

                state.RemoveCrmLink();
            }
        }

        /// <summary>
        /// Remove an outlook item and its associated sync state.
        /// </summary>
        /// <param name="syncState">The sync state of the item to remove.</param>
        protected void RemoveItemAndSyncState(SyncState<OutlookItemType> syncState)
        {
            this.LogItemAction(syncState.OutlookItem, "Synchroniser.RemoveItemAndSyncState, deleting item");
            try
            {
                syncState.DeleteItem();
            }
            catch (Exception ex)
            {
                Log.Error("Synchroniser.RemoveItemAndSyncState: Exception  oItem.oItem.Delete", ex);
            }
            this.RemoveItemSyncState(syncState);
        }

        /// <summary>
        /// Remove an item from ItemsSyncState.
        /// </summary>
        /// <param name="item">The sync state of the item to remove.</param>
        protected void RemoveItemSyncState(SyncState<OutlookItemType> item)
        {
            this.LogItemAction(item.OutlookItem, "Synchroniser.RemoveItemSyncState, removed item from queue");
            SyncStateManager.Instance.RemoveSyncState(item);
        }

        /// <summary>
        /// Given a list of items which exist in Outlook but are missing from CRM, resolve
        /// how to handle them.
        /// </summary>
        /// <param name="itemsToResolve">The list of items to resolve.</param>
        protected virtual void ResolveUnmatchedItems(IEnumerable<SyncState<OutlookItemType>> itemsToResolve)
        {
            foreach (var unresolved in itemsToResolve)
            {
                switch (unresolved.TxState)
                {
                    case TransmissionState.PendingDeletion:
                        /* If it's to resolve and marked pending deletion, we delete it
                         * (unresolved on two successive iterations): */
                        this.RemoveItemAndSyncState(unresolved);
                        break;

                    case TransmissionState.Synced:
                        if (unresolved.ExistedInCrm)
                        {
                            /* if it's unresolved but it used to exist in CRM, it's probably been deleted from
                             * CRM. Mark it pending deletion and check again next iteration. */
                            unresolved.SetPendingDeletion();
                        }
                        break;

                    case TransmissionState.Pending:
                    case TransmissionState.PresentAtStartup:
                        if (unresolved.ShouldSyncWithCrm)
                        {
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
                        }
                        break;

                    case TransmissionState.Queued:
                        if (unresolved.ShouldSyncWithCrm)
                        {
                            try
                            {
                                /* if it's queued and should be synced send it. */
                                AddOrUpdateItemFromOutlookToCrm(unresolved);
                            }
                            catch (BadStateTransition bst)
                            {
                                Log.Error($"Transition exception in ResolveUnmatchedItems", bst);
                            }
                        }
                        break;

                    default:
                        unresolved.SetPending();
                        break;
                }
            }

            foreach (SyncState resolved in SyncStateManager.Instance.GetSynchronisedItems<SyncStateType>()
                .Where(s => s.TxState == TransmissionState.PendingDeletion &&
                !itemsToResolve.Contains(s)))
            {
                /* finally, if there exists an item which had been marked pending deletion, but it has
                 *  been found in CRM (i.e. not in unresolved), mark it as synced */
                ((SyncState<OutlookItemType>)resolved).SetSynced();
            }
        }

        /// <summary>
        /// Save this item.
        /// </summary>
        /// <remarks>
        /// Because Outlook items are not proper objects, you cannot call the Save method of
        /// an Outlook item without knowing its exact class explicitly. So there are what look
        /// like redundant specialisations of this method; they aren't.
        /// </remarks>
        /// <param name="olItem">The item to save.</param>
        protected abstract void SaveItem(OutlookItemType olItem);

        /// <summary>
        /// Specialisations should return false if there's a good reason why we should
        /// NOT sync this item.
        /// </summary>
        /// <param name="folder">The folder to synchronise into.</param>
        /// <param name="crmType">The CRM type of the candidate item.</param>
        /// <param name="crmItem">The candidate item from CRM.</param>
        /// <returns>true</returns>
        protected virtual bool ShouldAddOrUpdateItemFromCrmToOutlook(Outlook.MAPIFolder folder, string crmType, EntryValue crmItem)
        {
            return true;
        }

        /// <summary>
        /// Perform all the necessary checking before adding or updating an item on CRM.
        /// </summary>
        /// <remarks>
        /// TODO TODO TODO: This does NOT actually do all the checking. Checking is also
        /// done in SyncState.ShouldSyncWithCRM, and possibly other places. Fix.
        /// </remarks>
        /// <param name="olItem">The item we may seek to add or update, presumed to be of
        /// my default item type.</param>
        /// <returns>true if we may attempt to add or update that item.</returns>
        protected virtual bool ShouldAddOrUpdateItemFromOutlookToCrm(OutlookItemType olItem)
        {
            bool result;
            string prefix = "Synchoniser.ShouldAddOrUpdateItemFromOutlookToCrm";

            try
            {
                if (olItem == null)
                {
                    Log.Warn($"{prefix}: attempt to send null {this.DefaultCrmModule}?");
                    result = false;
                }
                else
                {
                    if (SyncDirection.AllowOutbound(Direction))
                    {
                        if (this.permissionsCache.HasImportAccess(this.DefaultCrmModule))
                        {
                            if (this.GetSensitivity(olItem) == Outlook.OlSensitivity.olNormal)
                            {
                                result = true;
                            }
                            else
                            {
                                Log.Info($"{prefix}: {this.DefaultCrmModule} not added to CRM because its sensitivity is not public.");
                                result = false;
                            }
                        }
                        else
                        {
                            Log.Info($"{prefix}: {this.DefaultCrmModule} not added to CRM because import access is not granted.");
                            result = false;
                        }
                    }
                    else
                    {
                        Log.Info($"{prefix}: {this.DefaultCrmModule} not added to CRM because synchronisation is not enabled.");
                        result = false;
                    }
                }
            }
            catch (Exception any)
            {
                Log.Error($"{prefix}: unexpected failure while checking {this.DefaultCrmModule}.", any);
                result = false;
            }

            return result;
        }

        /// <summary>
        /// Should the item represented by this sync state be synchronised now?
        /// </summary>
        /// <param name="syncState">The sync state under consideration.</param>
        /// <returns>True if this synchroniser relates to the current tab and the timing logic is satisfied.</returns>
        protected bool ShouldPerformSyncNow(SyncState<OutlookItemType> syncState)
        {
            return (syncState.ShouldPerformSyncNow());
        }

        /// <summary>
        /// Add the magic 'SuiteCRM' category to the Outlook mapi namespace, if it does not
        /// already exist.
        /// </summary>
        private void AddSuiteCrmOutlookCategory()
        {
            Outlook.NameSpace oNS = this.Application.GetNamespace("mapi");
            if (oNS.Categories["SuiteCRM"] == null)
            {
                oNS.Categories.Add("SuiteCRM", Outlook.OlCategoryColor.olCategoryColorGreen,
                    Outlook.OlCategoryShortcutKey.olCategoryShortcutKeyNone);
            }
        }
    }
}
