using System;

namespace SuiteCRMAddIn.BusinessLogic
{
    using Outlook = Microsoft.Office.Interop.Outlook;

    public abstract class SyncState<ItemType>
    {
        public abstract string CrmType { get; }

        public string CrmEntryId { get; set; }

        public DateTime OModifiedDate { get; set; }

        public ItemType OutlookItem { get; set; }

        public int IsUpdate { get; set; }

        public bool ExistedInCrm => !string.IsNullOrEmpty(CrmEntryId);

        /// <summary>
        /// Precisely 'this.OutlookItem.EntryId'.
        /// </summary>
        /// <remarks>Outlook item classes do not inherit from a common base class, so generic client code cannot refer to 'OutlookItem.EntryId'.</remarks>
        public abstract string OutlookItemEntryId { get; }

        /// <summary>
        /// Precisely 'this.OutlookItem.Sensitivity'.
        /// </summary>
        /// <remarks>Outlook item classes do not inherit from a common base class, so generic client code cannot refer to 'OutlookItem.Sensitivity'.</remarks>
        public abstract Outlook.OlSensitivity OutlookItemSensitivity { get; }
    }
}
