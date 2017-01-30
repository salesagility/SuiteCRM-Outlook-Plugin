using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    public class AppointmentSyncState: SyncState<Outlook.AppointmentItem>
    {
        public string SType { get; set; }
    }
}
