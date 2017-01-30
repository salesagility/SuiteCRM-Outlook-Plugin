using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    public class ContactSyncState: SyncState<Outlook.ContactItem>
    {
        public override string CrmType => "Contacts";
    }
}
