using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    public class AppointmentSyncState: SyncState<Outlook.AppointmentItem>
    {
        public AppointmentSyncState(string crmType)
        {
            CrmType = crmType;
        }

        public override string CrmType { get; }

        public override string OutlookItemEntryId => OutlookItem.EntryID;

        public override Outlook.OlSensitivity OutlookItemSensitivity => OutlookItem.Sensitivity;

        public override Outlook.UserProperties OutlookUserProperties => OutlookItem.UserProperties;
    }
}
