using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    public class ContactSyncState: SyncState<Outlook.ContactItem>
    {
        public override string CrmType => ContactSyncing.CrmModule;

        public override bool ShouldSyncWithCrm => IsPublic;

        public override string OutlookItemEntryId => OutlookItem.EntryID;

        public override Outlook.OlSensitivity OutlookItemSensitivity => OutlookItem.Sensitivity;

        public override Outlook.UserProperties OutlookUserProperties => OutlookItem.UserProperties;

        /// <summary>
        /// Don't actually delete contact items from Outlook; instead, mark them private so they
        /// don't get copied back to CRM.
        /// </summary>
        public override void DeleteItem()
        {
            this.OutlookItem.Sensitivity = Microsoft.Office.Interop.Outlook.OlSensitivity.olPrivate;
        }
    }
}
