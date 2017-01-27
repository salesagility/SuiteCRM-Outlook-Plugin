using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    public class ContactSyncState
    {
        public string SEntryID { get; set; }

        public DateTime OModifiedDate { get; set; }

        public Outlook.ContactItem OutlookItem { get; set; }

        public bool Touched { get; set; }

        public bool Delete { get; set; }

        public int IsUpdate { get; set; }
    }
}
