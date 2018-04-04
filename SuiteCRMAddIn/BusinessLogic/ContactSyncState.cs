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
    using SuiteCRMAddIn.ProtoItems;
    using Extensions;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System.Runtime.InteropServices;
    using SuiteCRMClient.Logging;

    /// <summary>
    /// A SyncState for Contact items.
    /// </summary>
    public class ContactSyncState: SyncState<Outlook.ContactItem>
    {
        public ContactSyncState(Outlook.ContactItem oItem, string crmId, DateTime modified) : base(oItem, crmId, modified)
        {
        }

        public override string CrmType => ContactSyncing.CrmModule;

        public override bool ShouldSyncWithCrm => IsPublic;

        public override string OutlookItemEntryId => OutlookItem.EntryID;

        public override Outlook.OlSensitivity OutlookItemSensitivity => OutlookItem.Sensitivity;

        public override Outlook.UserProperties OutlookUserProperties => OutlookItem.UserProperties;

        public override string Description
        {
            get
            {
                Outlook.UserProperty olPropertyEntryId = olItem.UserProperties[Synchroniser<Outlook.ContactItem>.CrmIdPropertyName];
                string crmId = olPropertyEntryId == null ?
                    "[not present]" :
                    olPropertyEntryId.Value;
                return $"\tOutlook Id  : {olItem.EntryID}\n\tCRM Id      : {crmId}\n\tFull name   : '{olItem.FullName}'\n\tSensitivity : {olItem.Sensitivity}";
            }
        }


        /// <summary>
        /// Don't actually delete contact items from Outlook; instead, mark them private so they
        /// don't get copied back to CRM.
        /// </summary>
        public override void DeleteItem()
        {
            this.OutlookItem.Sensitivity = Microsoft.Office.Interop.Outlook.OlSensitivity.olPrivate;
        }

        /// <summary>
        /// Construct a JSON-serialisable representation of this contact item.
        /// </summary>
        internal override ProtoItem<Outlook.ContactItem> CreateProtoItem(Outlook.ContactItem outlookItem)
        {
            return new ProtoContact(outlookItem);
        }

        public override void RemoveSynchronisationProperties()
        {
            olItem.ClearSynchronisationProperties();
        }

    }
}
