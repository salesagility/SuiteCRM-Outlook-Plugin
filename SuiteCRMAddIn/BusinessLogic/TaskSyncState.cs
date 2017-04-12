using System;
using SuiteCRMAddIn.ProtoItems;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    public class TaskSyncState: SyncState<Outlook.TaskItem>
    {
        public override string CrmType => TaskSyncing.CrmModule;

        public override string OutlookItemEntryId => OutlookItem.EntryID;

        public override Outlook.OlSensitivity OutlookItemSensitivity => OutlookItem.Sensitivity;

        public override Outlook.UserProperties OutlookUserProperties => OutlookItem.UserProperties;

        public override void DeleteItem()
        {

            // this.OutlookItem.Delete();
        }

        internal override ProtoItem<Outlook.TaskItem> CreateProtoItem(Outlook.TaskItem outlookItem)
        {
            return new ProtoTask(outlookItem);
        }
    }
}
