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
    using Newtonsoft.Json;
    using Daemon;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Synchronise items of the class for which I am responsible.
    /// </summary>
    /// <typeparam name="OutlookItemType">The class of item for which I am responsible.</typeparam>
    public abstract class Synchroniser<OutlookItemType> : RepeatingProcess, IDisposable
        where OutlookItemType : class
    {
        /// <summary>
        /// The name of the modified date synchronisation property.
        /// </summary>
        protected const string ModifiedDatePropertyName = "SOModifiedDate";

        /// <summary>
        /// The name of the type synchronisation property.
        /// </summary>
        protected const string TypePropertyName = "SType";

        /// <summary>
        /// The name of the CRM ID synchronisation property.
        /// </summary>
        protected const string CrmIdPropertyName = "SEntryID";

        private readonly SyncContext context;

        /// <summary>
        /// A lock to prevent enqueueing the same new object twice in different
        /// threads (unlikely, since it should always be in the VSTA_main thread,
        /// but let's be paranoid).
        /// </summary>
        private object enqueueingLock = new object();

        /// <summary>
        /// The prefix for the fetch query, used in FetchRecordsFromCrm, q.v.
        /// </summary>
        protected string fetchQueryPrefix;

        // Keep a reference to the COM object on which we have event handlers, otherwise
        // when the reference is garbage-collected, the event-handlers are removed!
        private Outlook.Items _itemsCollection = null;

        private string _folderName;

        /// <summary>
        /// It appears that CRM sends us back strings HTML escaped.
        /// </summary>
        protected JsonSerializerSettings deserialiseSettings = new JsonSerializerSettings()
        {
            StringEscapeHandling = StringEscapeHandling.EscapeHtml
        };

        /// <summary>
        /// A cache for CRM premissions to prevent continually asking for them.
        /// </summary>
        protected readonly CRMPermissionsCache<OutlookItemType> permissionsCache;

        /// <summary>
        /// Construct a new instance of a synchroniser with this thread name and context.
        /// </summary>
        /// <param name="threadName">The name of the thread I shall create.</param>
        /// <param name="context">The context in which I shall work.</param>
        public Synchroniser(string threadName, SyncContext context) : base(threadName, context.Log)
        {
            this.context = context;
            InstallEventHandlers();
            this.AddSuiteCrmOutlookCategory();
            this.permissionsCache = new CRMPermissionsCache<OutlookItemType>(this, context.Log);
            GetOutlookItems(this.GetDefaultFolder());
        }


        /// <summary>
        /// Add the magis 'SuiteCRM' category to the Outlook mapi namespace, if it does not
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
                    Log.Debug($"{this.GetType().Name} SynchroniseAll starting");
                    this.SynchroniseAll();
                    Log.Debug($"{this.GetType().Name} SynchroniseAll completed");
                }
                else
                {
                    Log.Debug($"{this.GetType().Name}.SynchroniseAll not running because not enabled");
                }
            }
            else
            {
                Log.Debug($"{this.GetType().Name}.SynchroniseAll not running because no session");
            }
        }

        /// <summary>
        /// Run a single iteration of the synchronisation process for the items for which I am responsible.
        /// </summary>
        public virtual void SynchroniseAll()
        {
            if (this.permissionsCache.HasExportAccess())
            {
                Outlook.MAPIFolder folder = GetDefaultFolder();

                SyncFolder(folder, this.DefaultCrmModule);
            }
            else
            {
                Log.Debug($"{this.GetType().Name}.SynchroniseAll not synchronising {this.DefaultCrmModule} because export access is denied");
            }
        }

        protected abstract void GetOutlookItems(Outlook.MAPIFolder folder);

        protected abstract void SyncFolder(Outlook.MAPIFolder folder, string crmModule);

        /// <summary>
        /// The name of the default CRM module (record type) that this synchroniser synchronises.
        /// </summary>
        public abstract string DefaultCrmModule
        {
            get;
        }

        protected SyncContext Context => context;

        protected Outlook.Application Application => Context.Application;

        protected clsSettings settings => Context.settings;

        private bool readyToExit = false;

        /// <summary>
        /// List of the synchronisation state of all items which may require synchronisation.
        /// </summary>
        protected ThreadSafeList<SyncState<OutlookItemType>> ItemsSyncState { get; set; } = new ThreadSafeList<SyncState<OutlookItemType>>();

        /// <summary>
        /// The direction(s) in which I sync
        /// </summary>
        public abstract SyncDirection.Direction Direction { get; }

        /// <summary>
        /// We need to prevent two simultaneous transmissions of the same object, so it's probably unsafe
        /// to have two threads transmitting contact items at the same time. But there's no reason why
        /// we should not transmit contact items and task items at the same time, for example. So each
        /// Synchorniser subclass will have its own transmission lock.
        /// </summary>
        /// <returns>A transmission lock.</returns>
        protected abstract object TransmissionLock { get; }

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

        /// <summary>
        /// Check whether this synchroniser is allowed import access for its default CRM module.
        /// </summary>
        /// <returns>true if this synchroniser is allowed import access for its default CRM module.</returns>
        protected bool HasImportAccess()
        {
            return this.permissionsCache.HasImportAccess(this.DefaultCrmModule);
        }

        public override int PrepareShutdown()
        {
            if (!this.readyToExit)
            {
                /* remove event handlers when preparing shutdown, to prevent recursive issues */
                this.RemoveEventHandlers();

                /* brute force save everything. Not happy about this... */
                foreach (var syncState in this.ItemsSyncState)
                {
                    this.SaveItem(syncState.OutlookItem);
                }

                this.readyToExit = true;
            }
            return base.PrepareShutdown();
        }

        /// <summary>
        /// Given a list of items which exist in Outlook but are missing from CRM, resolve
        /// how to handle them.
        /// </summary>
        /// <param name="itemsToResolve">The list of items to resolve.</param>
        /// <param name="crmModule">The type of items to resolve.</param>
        protected void ResolveUnmatchedItems(IEnumerable<SyncState<OutlookItemType>> itemsToResolve, string crmModule)
        {
            var toDeleteFromOutlook = itemsToResolve.Where(a => a.ExistedInCrm && a.CrmType == crmModule).ToList();
            var toCreateOnCrmServer = itemsToResolve.Where(a => !a.ExistedInCrm && a.CrmType == crmModule).ToList();

            foreach (var syncState in toDeleteFromOutlook)
            {
                this.RemoveItemAndSyncState(syncState);
            }

            foreach (var syncState in toCreateOnCrmServer)
            {
                AddOrUpdateItemFromOutlookToCrm(syncState, this.DefaultCrmModule);
            }
        }

        /// <summary>
        /// Perform all the necessary checking before adding or updating an item on CRM.
        /// </summary>
        /// <param name="item">The item we may seek to add or update, presumed to be of
        /// my default item type.</param>
        /// <returns>true if we may attempt to add or update that item.</returns>
        protected bool ShouldAddOrUpdateItemFromOutlookToCrm(OutlookItemType item)
        {
            return this.ShouldAddOrUpdateItemFromOutlookToCrm(item, this.DefaultCrmModule);
        }

        /// <summary>
        /// Perform all the necessary checking before adding or updating an item on CRM.
        /// </summary>
        /// <remarks>
        /// TODO TODO TODO: This does NOT actually do all the checking. Checking is also
        /// done in SyncState.ShouldSyncWithCRM, and possibly other places. Fix.
        /// </remarks>
        /// <param name="item">The item we may seek to add or update.</param>
        /// <param name="crmType">The CRM type of that item.</param>
        /// <returns>true if we may attempt to add or update that item.</returns>
        protected bool ShouldAddOrUpdateItemFromOutlookToCrm(OutlookItemType item, string crmType)
        {
            bool result;
            string prefix = "Synchoniser.ShouldAddOrUpdateItemFromOutlookToCrm";

            try
            {
                if (item == null)
                {
                    Log.Warn($"{prefix}: attempt to send null {crmType}?");
                    result = false;
                }
                else
                {
                    if (SyncDirection.AllowOutbound(Direction))
                    {
                        if (this.permissionsCache.HasImportAccess(crmType))
                        {
                            if (this.GetSensitivity(item) == Outlook.OlSensitivity.olNormal)
                            {
                                result = true;
                            }
                            else
                            {
                                Log.Info($"{prefix}: {crmType} not added to CRM because its sensitivity is not public.");
                                result = false;
                            }
                        }
                        else
                        {
                            Log.Info($"{prefix}: {crmType} not added to CRM because import access is not granted.");
                            result = false;
                        }
                    }
                    else
                    {
                        Log.Info($"{prefix}: {crmType} not added to CRM because synchronisation is not enabled.");
                        result = false;
                    }
                }
            }
            catch (Exception any)
            {
                Log.Error($"{prefix}: unexpected failure while checking {crmType}.", any);
                result = false;
            }

            return result;
        }

        /// <summary>
        /// Return the sensitivity of this outlook item.
        /// </summary>
        /// <remarks>
        /// Outlook item classes do not inherit from a common base class, so generic client code cannot refer to 'OutlookItem.Sensitivity'.
        /// </remarks>
        /// <param name="item">The outlook item whose sensitivity is required.</param>
        /// <returns>the sensitivity of the item.</returns>
        internal abstract Outlook.OlSensitivity GetSensitivity(OutlookItemType item);

        /// <summary>
        /// Given a list of items which exist in Outlook but are missing from CRM, resolve
        /// how to handle them.
        /// </summary>
        /// <param name="itemsToResolve">The list of items to resolve.</param>
        /// <param name="crmModule">The type of items to resolve.</param>
        protected void ResolveUnmatchedItems(IEnumerable<SyncState<OutlookItemType>> itemsToResolve)
        {
            this.ResolveUnmatchedItems(itemsToResolve, DefaultCrmModule);
        }


        /// <summary>
        /// Add the item implied by this SyncState, which may not exist in CRM, to CRM.
        /// </summary>
        /// <param name="syncState">The sync state.</param>
        /// <returns>The id of the entry added or updated.</returns>
        internal virtual string AddOrUpdateItemFromOutlookToCrm(SyncState<OutlookItemType> syncState)
        {
            return this.AddOrUpdateItemFromOutlookToCrm(syncState, DefaultCrmModule);
        }


        /// <summary>
        /// Add the item implied by this SyncState, which may not exist in CRM, to CRM.
        /// </summary>
        /// <param name="syncState">The sync state.</param>
        /// <param name="crmType">The CRM type (name of CRM module) of the item.</param>
        /// <returns>The id of the entry added or updated.</returns>
        internal virtual string AddOrUpdateItemFromOutlookToCrm(SyncState<OutlookItemType> syncState, string crmType)
        {
            return this.AddOrUpdateItemFromOutlookToCrm(syncState, crmType, syncState.CrmEntryId);
        }


        /// <summary>
        /// Add the Outlook item referenced by this sync state, which may not exist in CRM, to CRM.
        /// </summary>
        /// <param name="syncState">The sync state referencing the outlook item to add.</param>
        /// <param name="crmType">The CRM type ('module') to which it should be added</param>
        /// <param name="entryId">The id of this item in CRM, if known (in which case I should be doing
        /// an update, not an add).</param>
        /// <returns>The id of the entry added o</returns>
        internal virtual string AddOrUpdateItemFromOutlookToCrm(SyncState<OutlookItemType> syncState, string crmType, string entryId = "")
        {
            string result = entryId;

            if (this.ShouldAddOrUpdateItemFromOutlookToCrm(syncState.OutlookItem, crmType))
            {
                OutlookItemType outlookItem = syncState.OutlookItem;

                try
                {
                    lock (this.TransmissionLock)
                    {
                        LogItemAction(outlookItem, "Synchroniser.AddOrUpdateItemFromOutlookToCrm, Despatching");

                        if (syncState != null)
                        {
                            syncState.SetTransmitted();
                        }

                        result = ConstructAndDespatchCrmItem(outlookItem, crmType, entryId);
                        if (!string.IsNullOrEmpty(result))
                        {
                            var utcNow = DateTime.UtcNow;
                            EnsureSynchronisationPropertiesForOutlookItem(outlookItem, utcNow.ToString(), crmType, result);
                            this.SaveItem(outlookItem);

                            syncState.SetSynced(result);
                        }
                        else
                        {
                            Log.Warn("AppointmentSyncing.AddItemFromOutlookToCrm: Invalid CRM Id returned; item may not be stored.");
                            syncState.SetPending();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("Synchroniser.AddOrUpdateItemFromOutlookToCrm", ex);
                    syncState.SetPending();
                }
                finally
                {
                    this.SaveItem(outlookItem);
                }
            }

            return result;
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
        /// Find the SyncState whose item is this item; if it does not already exist, construct and return it.
        /// </summary>
        /// <param name="oItem">The item to find.</param>
        /// <returns>the SyncState whose item is this item</returns>
        protected SyncState<OutlookItemType> AddOrGetSyncState(OutlookItemType oItem)
        {
            var existingState = GetExistingSyncState(oItem);
            if (existingState != null)
            {
                if (existingState.OutlookItem != oItem)
                {
                    /* if Outlook only holds one item with the same id, then this line MUST be redundant.
                     * TODO: check logs and simplify this logic if the issue does not occur */
                    existingState.OutlookItem = oItem;
                    Log.Error($"Should never happen - two Outlook items with same id ({GetOutlookEntryId(oItem)})?");
                }
                return existingState;
            }
            else
            {
                return ConstructAndAddSyncState(oItem);
            }
        }

        /// <summary>
        /// Constructs a new SyncState object for this Outlook item and adds it to my 
        /// collection of sync states.
        /// </summary>
        /// <param name="oItem">The Outlook item to wrap</param>
        /// <returns>The sync state added.</returns>
        private SyncState<OutlookItemType> ConstructAndAddSyncState(OutlookItemType oItem)
        {
            SyncState<OutlookItemType> newState = ConstructSyncState(oItem);
            ItemsSyncState.Add(newState);
            return newState;
        }

        /// <summary>
        /// Construct and return a new sync state representing this item.
        /// </summary>
        /// <param name="oItem">The item</param>
        /// <returns>a new sync state representing this item.</returns>
        protected abstract SyncState<OutlookItemType> ConstructSyncState(OutlookItemType oItem);

        /// <summary>
        /// Get the existing sync state representing this item, if it exists, else null.
        /// </summary>
        /// <param name="oItem">The item</param>
        /// <returns>the existing sync state representing this item, if it exists, else null.</returns>
        protected SyncState<OutlookItemType> GetExistingSyncState(OutlookItemType olItem)
        {
            SyncState<OutlookItemType> result;

            if (olItem == null)
            {
                result = null;
            }
            else
            {
                var olItemEntryId = GetOutlookEntryId(olItem);
                try
                {
                    /* if there are duplicate entries I want them logged */
                    result = this.ItemsSyncState.Where(a => a.OutlookItem != null)
                        .Where(a => !string.IsNullOrEmpty(this.GetOutlookEntryId(a.OutlookItem)))
                        .Where(a => !a.IsDeletedInOutlook)
                        .SingleOrDefault(a => this.GetOutlookEntryId(a.OutlookItem).Equals(olItemEntryId));
                }
                catch (InvalidOperationException notUnique)
                {
                    Log.Error(
                        String.Format(
                            "AppointmentSyncing.AddItemFromOutlookToCrm: Outlook Id {0} was not unique in this.ItemsSyncState?",
                            olItemEntryId),
                        notUnique);

                    /* but if it isn't unique, the first will actually do for now */
                    result = this.ItemsSyncState.Where(a => a.OutlookItem != null)
                        .Where(a => !string.IsNullOrEmpty(this.GetOutlookEntryId(a.OutlookItem)))
                        .Where(a => !a.IsDeletedInOutlook)
                        .FirstOrDefault(a => this.GetOutlookEntryId(a.OutlookItem).Equals(olItemEntryId));
                }
            }

            return result;
        }

        /// <summary>
        /// Get the existing sync state representing this item, if it exists, else null.
        /// </summary>
        /// <param name="crmItem">The item</param>
        /// <returns>the existing sync state representing this item, if it exists, else null.</returns>
        protected SyncState<OutlookItemType> GetExistingSyncState(eEntryValue crmItem)
        {
            return crmItem == null ?
                null :
                this.GetExistingSyncState(crmItem.GetValueAsString("id"));
        }

        /// <summary>
        /// Get the existing sync state representing the item with this CRM id, if it exists, else null.
        /// </summary>
        /// <param name="crmItemId">The id of a CRM item</param>
        /// <returns>the existing sync state representing the item with this CRM id, if it exists, else null.</returns>
        protected SyncState<OutlookItemType> GetExistingSyncState(string crmItemId)
        {
            SyncState<OutlookItemType> result;
            try
            {
                /* if there are duplicate entries I want them logged */
                result = ItemsSyncState.SingleOrDefault(a => a.CrmEntryId == crmItemId);
            }
            catch (InvalidOperationException notUnique)
            {
                Log.Error(
                    String.Format(
                        "AppointmentSyncing.AddItemFromOutlookToCrm: CRM Id {0} was not unique in this.ItemsSyncState?",
                        crmItemId),
                    notUnique);

                /* but if it isn't unique, the first will actually do for now */
                result = ItemsSyncState.FirstOrDefault(a => a.CrmEntryId == crmItemId);
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
        /// Find the SyncState whose item is this item; if it does not already exist, construct and return it.
        /// </summary>
        /// <param name="oItem">The item to find.</param>
        /// <param name="modified">The modified time to set.</param>
        /// <param name="crmId">The id of this item in CRM.</param>
        /// <returns>the SyncState whose item is this item</returns>
        protected SyncState<OutlookItemType> AddOrGetSyncState(OutlookItemType oItem, DateTime modified, string crmId)
        {
            var result = this.AddOrGetSyncState(oItem);
            result.OModifiedDate = DateTime.UtcNow;
            result.CrmEntryId = crmId;

            return result;
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
        /// <param name="crmType">The type within CRM to which the item should be added.</param>
        /// <param name="entryId">The corresponding entry id in CRM, if known.</param>
        /// <returns>The CRM id of the object created or modified.</returns>
        protected abstract string ConstructAndDespatchCrmItem(OutlookItemType olItem, string crmType, string entryId);

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
            EnsureSynchronisationPropertyForOutlookItem(olItem, ModifiedDatePropertyName, modifiedDate);
            EnsureSynchronisationPropertyForOutlookItem(olItem, TypePropertyName, type);
            EnsureSynchronisationPropertyForOutlookItem(olItem, CrmIdPropertyName, entryId);
        }

        /// <summary>
        /// Set up synchronisation properties for this outlook item from this CRM item, assuming my default CRM module.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmItem">The CRM item.</param>
        protected virtual void EnsureSynchronisationPropertiesForOutlookItem(OutlookItemType olItem, eEntryValue crmItem)
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
        protected virtual void EnsureSynchronisationPropertiesForOutlookItem(OutlookItemType olItem, eEntryValue crmItem, string type)
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
        /// Returns true iff user is currently focussed on this (Contacts/Appointments/Tasks) tab.
        /// </summary>
        /// <remarks>
        /// This is used in determining whether an item is in fact newly created by the user;
        /// it has a certain code smell to it.
        /// </remarks>
        protected abstract bool IsCurrentView { get; }

        /// <summary>
        /// Returns true iff local (Outlook) deletions should be propagated to the server.
        /// </summary>
        /// <remarks>TODO: Why should this ever be false?</remarks>
        protected abstract bool PropagatesLocalDeletions { get; }

        /// <summary>
        /// Deal, in CRM, with items deleted in Outlook.
        /// </summary>
        protected void RemoveDeletedItems()
        {
            if (IsCurrentView && PropagatesLocalDeletions)
            {
                // Make a copy of the list to avoid mutation error while iterating:
                var syncStatesCopy = new List<SyncState<OutlookItemType>>(ItemsSyncState);
                foreach (var oItem in syncStatesCopy)
                {
                    var shouldDeleteFromCrm = oItem.IsDeletedInOutlook || !oItem.ShouldSyncWithCrm;
                    if (shouldDeleteFromCrm) RemoveFromCrm(oItem);
                    if (oItem.IsDeletedInOutlook) ItemsSyncState.Remove(oItem);
                }
            }
            else
            {
                var items = ItemsSyncState.Where(x => x.IsDeletedInOutlook).Count();
                if (items > 0)
                {
                    Log.Error($"Possibly bug #95: was attempting to delete {items} items from CRM");
                }
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
                    eNameValue[] data = new eNameValue[2];
                    data[0] = clsSuiteCRMHelper.SetNameValuePair("id", crmEntryId);
                    data[1] = clsSuiteCRMHelper.SetNameValuePair("deleted", "1");
                    clsSuiteCRMHelper.SetEntryUnsafe(data, state.CrmType);
                }

                state.RemoveCrmLink();
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
        /// Fetch records in pages from CRM, and merge them into Outlook.
        /// </summary>
        /// <param name="folder">The folder to be synchronised.</param>
        /// <param name="crmModule">The name of the CRM module to synchronise with.</param>
        /// <param name="untouched">A list of all known Outlook items, from which those modified by this method are removed.</param>
        protected virtual void MergeRecordsFromCrm(Outlook.MAPIFolder folder, string crmModule, HashSet<SyncState<OutlookItemType>> untouched)
        {
            int thisOffset = 0; // offset of current page of entries
            int nextOffset = 0; // offset of the next page of entries, if any.

            /* get candidates for syncrhonisation from SuiteCRM one page at a time */
            do
            {
                /* update the offset to the offset of the next page */
                thisOffset = nextOffset;

                /* fetch the page of entries starting at thisOffset */
                eGetEntryListResult entriesPage = clsSuiteCRMHelper.GetEntryList(crmModule,
                    String.Format(fetchQueryPrefix, clsSuiteCRMHelper.GetUserId()),
                    0, "date_start DESC", thisOffset, false,
                    clsSuiteCRMHelper.GetSugarFields(crmModule));

                /* get the offset of the next page */
                nextOffset = entriesPage.next_offset;

                /* when there are no more entries, we'll get a zero-length entry list and nextOffset
                 * will have the same value as thisOffset */
                AddOrUpdateItemsFromCrmToOutlook(entriesPage.entry_list, folder, untouched, crmModule);
            }
            while (thisOffset != nextOffset);

        }


        /// <summary>
        /// Update these items, which may or may not already exist in Outlook.
        /// </summary>
        /// <param name="items">The items to be synchronised.</param>
        /// <param name="folder">The outlook folder to synchronise into.</param>
        /// <param name="untouched">A list of sync states of existing items which have 
        /// not yet been synchronised; this list is modified (destructuvely changed) 
        /// by the action of this method.</param>
        /// <param name="crmType">The CRM record type ('module') to be fetched.</param>
        protected virtual void AddOrUpdateItemsFromCrmToOutlook(
            eEntryValue[] items,
            Outlook.MAPIFolder folder,
            HashSet<SyncState<OutlookItemType>> untouched,
            string crmType)
        {
            foreach (var item in items)
            {
                try
                {
                    var state = AddOrUpdateItemFromCrmToOutlook(folder, crmType, item);
                    if (state != null)
                    {
                        // i.e., the entry was updated...
                        untouched.Remove(state);
                        state.SetSynced();
                        LogItemAction(state.OutlookItem, "Synchroniser.UpdateItemsFromCrmToOutlook, item removed from untouched");
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("Synchroniser.UpdateItemsFromCrmToOutlook", ex);
                }
            }
        }

        /// <summary>
        /// Update a single item in the specified Outlook folder with changes from CRM. If the item
        /// does not exist, create it.
        /// </summary>
        /// <param name="folder">The folder to synchronise into.</param>
        /// <param name="crmType">The CRM type of the candidate item.</param>
        /// <param name="candidateItem">The candidate item from CRM.</param>
        /// <returns>The synchronisation state of the item updated (if it was updated).</returns>
        protected abstract SyncState<OutlookItemType> AddOrUpdateItemFromCrmToOutlook(Outlook.MAPIFolder folder, string crmType, eEntryValue candidateItem);

        /// <summary>
        /// Log a message regarding this Outlook item, with detail of the item.
        /// </summary>
        /// <param name="olItem">The outlook item.</param>
        /// <param name="message">The message to be logged.</param>
        protected abstract void LogItemAction(OutlookItemType olItem, string message);


        public void Dispose()
        {
            // RemoveEventHandlers();
        }

        protected void InstallEventHandlers()
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

        private void RemoveEventHandlers()
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

        protected void Items_ItemAdd(object outlookItem)
        {
            Log.Warn($"Outlook {_folderName} ItemAdd");
            try
            {
                OutlookItemAdded(outlookItem as OutlookItemType);
            }
            catch (Exception problem)
            {
                Log.Error($"{_folderName} ItemAdd failed", problem);
            }
        }

        protected void Items_ItemChange(object outlookItem)
        {
            Log.Debug($"Outlook {_folderName} ItemChange");
            try
            {
                OutlookItemChanged(outlookItem as OutlookItemType);
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
        /// Entry point from event handler when an item is added in Outlook.
        /// </summary>
        /// <remarks>Should always run in the 'VSTA_main' thread.</remarks>
        /// <param name="olItem">The item that has been added.</param>
        protected virtual void OutlookItemAdded(OutlookItemType olItem)
        {
            LogItemAction(olItem, "AppointmentSyncing.OutlookItemAdded");

            if (olItem != null)
            {
                lock (enqueueingLock)
                {
                    if (IsCurrentView && this.GetExistingSyncState(olItem) == null)
                    {
                        SyncState<OutlookItemType> state = this.ConstructAndAddSyncState(olItem);
                        DaemonWorker.Instance.AddTask(new TransmitNewAction<OutlookItemType>(this, state, this.DefaultCrmModule));
                    }
                    else
                    {
                        Log.Warn($"AppointmentSyncing.OutlookItemAdded: item {this.GetOutlookEntryId(olItem)} had already been added");
                    }
                }
            }
        }


        /// <summary>
        /// Entry point from event handler, called when an Outlook item of class AppointmentItem 
        /// is believed to have changed.
        /// </summary>
        /// <param name="olItem">The item which has changed.</param>
        protected void OutlookItemChanged(OutlookItemType olItem)
        {
            LogItemAction(olItem, "Syncroniser.OutlookItemChanged");
            try
            {
                var syncStateForItem = GetExistingSyncState(olItem);
                if (syncStateForItem != null)
                {
                    if (this.ShouldPerformSyncNow(syncStateForItem))
                    {
                        DaemonWorker.Instance.AddTask(new TransmitUpdateAction<OutlookItemType>(this, syncStateForItem));
                    }
                    else if (!syncStateForItem.ShouldSyncWithCrm)
                    {
                        this.RemoveFromCrm(syncStateForItem);
                    }
                }
                else
                {
                    /* we don't have a sync state for this item (presumably formerly private);
                     * that's OK, treat it as new */
                    OutlookItemAdded(olItem);
                }
            }
            finally
            {
                this.SaveItem(olItem);
            }
        }

        public abstract Outlook.MAPIFolder GetDefaultFolder();


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
            this.ItemsSyncState.Remove(item);
        }

        /// <summary>
        /// Should the item represented by this sync state be synchronised now?
        /// </summary>
        /// <param name="syncState">The sync state under consideration.</param>
        /// <returns>True if this synchroniser relates to the current tab and the timing logic is satisfied.</returns>
        protected bool ShouldPerformSyncNow(SyncState<OutlookItemType> syncState)
        {
            return (IsCurrentView && syncState.ShouldPerformSyncNow());
        }
    }
}
