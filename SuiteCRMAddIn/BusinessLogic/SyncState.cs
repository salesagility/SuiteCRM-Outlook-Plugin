using System;

namespace SuiteCRMAddIn.BusinessLogic
{
    public abstract class SyncState<ItemType>
    {
        public abstract string CrmType { get; }

        public string CrmEntryId { get; set; }

        public DateTime OModifiedDate { get; set; }

        public ItemType OutlookItem { get; set; }

        public bool Touched { get; set; }

        public bool Delete { get; set; }

        public int IsUpdate { get; set; }
    }
}
