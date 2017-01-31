using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    using System.Collections.Generic;
    using System.Linq;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System.Runtime.InteropServices;

    public abstract class Syncing<OutlookItemType>
    {
        private readonly SyncContext _context;

        public Syncing(SyncContext context)
        {
            _context = context;
        }

        protected SyncContext Context => _context;

        protected Outlook.Application Application => Context.Application;

        protected clsSettings settings => Context.settings;

        protected ILogger Log => Context.Log;

        protected List<SyncState<OutlookItemType>> ItemsSyncState { get; set; } = null;

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

        private static void RemoveFromCrm(SyncState<OutlookItemType> oItem)
        {
            var crmEntryId = oItem.CrmEntryId;
            if (!string.IsNullOrEmpty(crmEntryId))
            {
                eNameValue[] data = new eNameValue[2];
                data[0] = clsSuiteCRMHelper.SetNameValuePair("id", crmEntryId);
                data[1] = clsSuiteCRMHelper.SetNameValuePair("deleted", "1");
                clsSuiteCRMHelper.SetEntryUnsafe(data, oItem.CrmType);
            }
        }
    }
}
