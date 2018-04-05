/**
 * Outlook integration for SuiteCRM.
 * @package Outlook integration for SuiteCRM
 * @copyright SalesAgility Ltd http://www.salesagility.com
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU LESSER GENERAL PUBLIC LICENCE as published by
 * the Free Software Foundation; either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU LESSER GENERAL PUBLIC LICENCE
 * along with this program; if not, see http://www.gnu.org/licenses
 * or write to the Free Software Foundation,Inc., 51 Franklin Street,
 * Fifth Floor, Boston, MA 02110-1301  USA
 *
 * @author SalesAgility <info@salesagility.com>
 */
namespace SuiteCRMAddIn.BusinessLogic
{
    using System;
    using ProtoItems;
    using Extensions;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System.Runtime.InteropServices;

    /// <summary>
    /// A SyncState for Contact items.
    /// </summary>
    public class TaskSyncState: SyncState<Outlook.TaskItem>
    {
        public TaskSyncState(Outlook.TaskItem item, string crmId, DateTime modifiedDate) : base(item, crmId, modifiedDate)
        {
        }

        public override string CrmType => TaskSyncing.CrmModule;

        public override string OutlookItemEntryId => OutlookItem.EntryID;

        public override Outlook.OlSensitivity OutlookItemSensitivity => OutlookItem.Sensitivity;

        public override Outlook.UserProperties OutlookUserProperties => OutlookItem.UserProperties;

        public override string Description
        {
            get
            {
                if (OutlookItem == null)
                {
                    return "[OutlookItem not set]";
                }
                else
                {
                    Outlook.UserProperty olPropertyEntryId = OutlookItem.UserProperties[Synchroniser<Outlook.TaskItem>.CrmIdPropertyName];
                    string crmId = olPropertyEntryId == null ?
                        "[not present]" :
                        olPropertyEntryId.Value;
                    return $"\tOutlook Id  : {OutlookItem.EntryID}\n\tCRM Id      : {crmId}\n\tSubject    : '{OutlookItem.Subject}'\n\tStatus      : {OutlookItem.Status}";
                }
            }
        }


        public override void DeleteItem()
        {
            this.OutlookItem.Delete();
        }

        /// <summary>
        /// Construct a JSON-serialisable representation of my task item.
        /// </summary>
        internal override ProtoItem<Outlook.TaskItem> CreateProtoItem(Outlook.TaskItem outlookItem)
        {
            return new ProtoTask(outlookItem);
        }

        public override void RemoveSynchronisationProperties()
        {
            OutlookItem.ClearSynchronisationProperties();
        }
    }
}
