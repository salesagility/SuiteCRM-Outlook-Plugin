using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using SuiteCRMClient;
using SuiteCRMClient.Logging;
using SuiteCRMClient.RESTObjects;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    // TODO: Make syncing thread-safe.
    public abstract class Syncing<OutlookItemType>: IDisposable
        where OutlookItemType: class
    {
        private readonly SyncContext _context;

        // Keep a reference to the COM object on which we have event handlers, otherwise
        // when the reference is garbage-collected, the event-handlers are removed!
        private Outlook.Items _itemsCollection = null;

        private string _folderName;

        public Syncing(SyncContext context)
        {
            _context = context;
            InstallEventHandlers();
        }

        protected SyncContext Context => _context;

        protected Outlook.Application Application => Context.Application;

        protected clsSettings settings => Context.settings;

        protected ILogger Log => Context.Log;

        protected List<SyncState<OutlookItemType>> ItemsSyncState { get; set; } = null;

        public abstract bool SyncingEnabled { get; }

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

        protected bool HasAccess(string moduleName, string permission)
        {
            try
            {
                eModuleList oList = clsSuiteCRMHelper.GetModules();
                return oList.modules1.FirstOrDefault(a => a.module_label == moduleName)
                    ?.module_acls1.FirstOrDefault(b => b.action == permission)
                    ?.access ?? false;
            }
            catch (Exception)
            {
                Log.Warn($"Cannot detect access {moduleName}/{permission}");
                return false;
            }
        }

        /// <summary>
        /// Returns true iif user is currently focussed on this (Contacts/Appointments/Tasks) tab.
        /// </summary>
        /// <remarks>TODO: Why should this make a difference?</remarks>
        protected abstract bool IsCurrentView { get; }

        /// <summary>
        /// Returns true iff local (Outlook) deletions should be propagated to the server.
        /// </summary>
        /// <remarks>TODO: Why should this ever be false?</remarks>
        protected abstract bool PropagatesLocalDeletions { get; }

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
        }

        protected void RemoveFromCrm(SyncState state)
        {
            if (!SyncingEnabled) return;
            var crmEntryId = state.CrmEntryId;
            if (!string.IsNullOrEmpty(crmEntryId))
            {
                eNameValue[] data = new eNameValue[2];
                data[0] = clsSuiteCRMHelper.SetNameValuePair("id", crmEntryId);
                data[1] = clsSuiteCRMHelper.SetNameValuePair("deleted", "1");
                clsSuiteCRMHelper.SetEntryUnsafe(data, state.CrmType);
            }

            state.RemoveCrmLink();
        }

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

        public void Dispose()
        {
            RemoveEventHandlers();
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

        protected abstract void OutlookItemAdded(OutlookItemType outlookItem);

        protected abstract void OutlookItemChanged(OutlookItemType outlookItem);

        public abstract Outlook.MAPIFolder GetDefaultFolder();
    }
}
