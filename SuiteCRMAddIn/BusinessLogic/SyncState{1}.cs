namespace SuiteCRMAddIn.BusinessLogic
{

    public abstract class SyncState<ItemType>: SyncState
    {
        public ItemType OutlookItem { get; set; }
    }
}
