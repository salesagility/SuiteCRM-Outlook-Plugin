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

        protected abstract bool IsCurrentView { get; }

        /// <summary>
        /// Returns true iff local (Outlook) deletions should be propagated to the server.
        /// </summary>
        protected abstract bool PropagatesLocalDeletions { get; }

        protected void RemoveDeletedItems(bool checkItemSensitivity)
        {
            if (IsCurrentView && PropagatesLocalDeletions)
            {
                var toBeDeleted = new HashSet<SyncState<OutlookItemType>>();
                foreach (var oItem in ItemsSyncState)
                {
                    try
                    {
                        // Has the side-effect of throwing an exception if the item has been deleted:
                        if (checkItemSensitivity && oItem.OutlookItemSensitivity != Outlook.OlSensitivity.olNormal)
                            continue;
                        var sID = oItem.OutlookItemEntryId;
                    }
                    catch (COMException)
                    {
                        eNameValue[] data = new eNameValue[2];
                        data[0] = clsSuiteCRMHelper.SetNameValuePair("id", oItem.CrmEntryId);
                        data[1] = clsSuiteCRMHelper.SetNameValuePair("deleted", "1");
                        clsSuiteCRMHelper.SetEntryUnsafe(data, oItem.CrmType);
                        toBeDeleted.Add(oItem);
                    }
                }
                ItemsSyncState.RemoveAll(a => toBeDeleted.Contains(a));
            }
        }
    }
}
