using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    public class AppointmentSyncState
    {
        public string SEntryID { get; set; }

        public string SType { get; set; }

        public DateTime OModifiedDate { get; set; }

        public Outlook.AppointmentItem OutlookItem { get; set; }

        public bool Touched { get; set; }

        public bool Delete { get; set; }

        public int IsUpdate { get; set; }
    }
}
