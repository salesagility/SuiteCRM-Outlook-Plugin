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

using System.Threading;
using SuiteCRMAddIn.Exceptions;

namespace SuiteCRMAddIn.BusinessLogic
{
    using System;
    using SuiteCRMAddIn.ProtoItems;
    using Extensions;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System.Runtime.InteropServices;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using SuiteCRMClient;

    /// <summary>
    /// A SyncState for Contact items.
    /// </summary>
    public class ContactSyncState: SyncState<Outlook.ContactItem>
    {
        private ILogger Log = Globals.ThisAddIn.Log;
        public ContactSyncState(Outlook.ContactItem oItem, CrmId crmId, DateTime modified) : base(oItem, crmId, modified)
        {
        }

        public override Outlook.Folder DefaultFolder => (Outlook.Folder)MapiNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);

        public override Outlook.ContactItem OutlookItem
        {
            get
            {
                var result = base.OutlookItem;

                try
                {
                    var check = result.EntryID;
                }
                catch (Exception ex) when (ex is InvalidComObjectException || ex is COMException)
                {
                    /* this is thrown if the reference for the item is not good in this thread */
                    object r = null;

                    for (var i = 0; r == null && i < 10; i++)
                    {
                        try
                        {
                            r = Globals.ThisAddIn.Application.Session.GetItemFromID(this.outlookItemId);
                        }
                        catch (Exception any)
                        {
                            Log.Error($"Failed to open item with id {this.outlookItemId} at attempt {i}", any);
                            Thread.Sleep(10000);
                        }
                    }

                    if (r is Outlook.ContactItem)
                    {
                        result = r as Outlook.ContactItem;
                        base.OutlookItem = result;
                    }
                    else
                    {
                        throw new MissingOutlookItemException(this.outlookItemId);
                    }
                }

                return result;
            }
        }

        /// <summary>
        /// If transmission was successful, clear the manual override if set.
        /// </summary>
        internal override void SetTransmitted()
        {
            base.SetTransmitted();
            this.OutlookItem.ClearManualOverride();
        }

        /// <summary>
        /// True if the Outlook item wrapped by this state may be synchronised even when synchronisation is set to none.
        /// </summary>
        public override bool IsManualOverride => this.OutlookItem.IsManualOverride();

        public override string CrmType => ContactSynchroniser.CrmModule;

        public override bool ShouldSyncWithCrm => IsPublic;

        public override string OutlookItemEntryId => OutlookItem.EntryID;

        public override Outlook.OlSensitivity OutlookItemSensitivity => OutlookItem.Sensitivity;

        public override Outlook.UserProperties OutlookUserProperties => OutlookItem.UserProperties;

        public override string Description
        {
            get
            {
                CrmId crmId = OutlookItem.GetCrmId();
                if (CrmId.IsInvalid(crmId)) { crmId = CrmId.Empty; }

                return $"\tOutlook Id  : {OutlookItem.EntryID}\n\tCRM Id      : {crmId}\n\tFull name   : '{OutlookItem.FullName}'\n\tSensitivity : {OutlookItem.Sensitivity}";
            }
        }

        public override string IdentifyingFields
        {
            get
            {
                return $"name: '{OutlookItem.FullName}'; email: '{OutlookItem.Email1Address}'";
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
            OutlookItem.ClearSynchronisationProperties();
        }


        /// <summary>
        /// Get a string representing the values of the distinct fields of this crmItem, 
        /// as a final fallback for identifying an otherwise unidentifiable object.
        /// </summary>
        /// <param name="crmItem">An item received from CRM.</param>
        /// <returns>An identifying string.</returns>
        /// <see cref="SyncState{ItemType}.IdentifyingFields"/> 
        internal static string GetDistinctFields(EntryValue crmItem)
        {
            // TODO: fix
            return $"subject: '{crmItem.GetValueAsString("name")}'; start: '{crmItem.GetValueAsDateTime("date_start")}'";
        }

        internal override void SaveItem()
        {
            this.OutlookItem?.Save();
        }

        protected override void CacheOulookItemId(Outlook.ContactItem olItem)
        {
            this.outlookItemId = olItem.EntryID;
        }
    }
}
