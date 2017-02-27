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
        private readonly SyncContext context;

        // Keep a reference to the COM object on which we have event handlers, otherwise
        // when the reference is garbage-collected, the event-handlers are removed!
        private Outlook.Items _itemsCollection = null;

        private string _folderName;

        public Synchroniser(string name, SyncContext context) : base(name, context.Log)
        {
            this.context = context;
            InstallEventHandlers();
        }

        /// <summary>
        /// If I am currently configured to do so, synchronise the items for which I am
        /// responsible once.
        /// </summary>
        internal override void PerformIteration()
        {
            if (Globals.ThisAddIn.HasCrmUserSession)
            {
                if (this.SyncingEnabled)
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
        public abstract void SynchroniseAll();

        protected SyncContext Context => context;

        protected Outlook.Application Application => Context.Application;

        protected clsSettings settings => Context.settings;


        /// <summary>
        /// List of the synchronisation state of all items which may require synchronisation.
        /// Note that this list is NOT thread safe. TODO: Reimplement using Thread-Safe
        /// Collections, probably ConcurrentBag. See
        /// https://msdn.microsoft.com/en-us/library/dd997305(v=vs.110).aspx
        /// </summary>
        protected ThreadSafeList<SyncState<OutlookItemType>> ItemsSyncState { get; set; } = null;

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
