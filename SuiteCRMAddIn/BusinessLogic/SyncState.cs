using System;

namespace SuiteCRMAddIn.BusinessLogic
{
    public abstract class SyncState<ItemType>
    {
        public string SEntryID { get; set; }

        public DateTime OModifiedDate { get; set; }

        public ItemType OutlookItem { get; set; }

        public bool Touched { get; set; }

        public bool Delete { get; set; }

        public int IsUpdate { get; set; }
    }
}
