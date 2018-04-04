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
    using Extensions;
    using ProtoItems;
    using System.Linq;
    using System.Text;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using SuiteCRMClient.Logging;

    public class AppointmentSyncState: SyncState<Outlook.AppointmentItem>
    {
        public AppointmentSyncState(Outlook.AppointmentItem item, string crmId, DateTime modifiedDate) : base(item, crmId, modifiedDate)
        {
        }

        /// <summary>
        /// When we're asked for the CrmType the underlying object may have ceased to
        /// exist - so cache it!
        /// </summary>
        private string crmType;


        /// <summary>
        /// The CRM type of the item I represent.
        /// </summary>
        public override string CrmType
        {
            get
            {
                try
                {
                    switch (olItem.MeetingStatus)
                    {
                        case Outlook.OlMeetingStatus.olNonMeeting:
                            crmType = AppointmentSyncing.AltCrmModule;
                            break;
                        default:
                            crmType = AppointmentSyncing.CrmModule;
                            break;
                    }
                    return crmType;
                }
                catch (COMException)
                {
                    return crmType;
                }
                catch (Exception)
                {
                    return string.Empty;
                }
            }
        }

        public override string Description
        {
            get
            {
                Outlook.UserProperty olPropertyEntryId = olItem.UserProperties[AppointmentSyncing.CrmIdPropertyName];
                string crmId = olPropertyEntryId == null ?
                    "[not present]" :
                    olPropertyEntryId.Value;
                StringBuilder bob = new StringBuilder();
                bob.Append($"\tOutlook Id  : {olItem.EntryID}\n\tCRM Id      : {crmId}\n\tSubject     : '{olItem.Subject}'\n\tSensitivity : {olItem.Sensitivity}\n\tStatus     : {olItem.MeetingStatus}\n\tRecipients:\n");
                foreach (Outlook.Recipient recipient in olItem.Recipients)
                {
                    bob.Append($"\t\t{recipient.Name}: {recipient.GetSmtpAddress()} - ({recipient.MeetingResponseStatus})\n");
                }

                return bob.ToString();
            }
        }

        public override string OutlookItemEntryId => OutlookItem.EntryID;

        public override Outlook.OlSensitivity OutlookItemSensitivity => OutlookItem.Sensitivity;

        public override Outlook.UserProperties OutlookUserProperties => OutlookItem.UserProperties;

        public override void DeleteItem()
        {
            this.OutlookItem.Delete();
        }

        /// <summary>
        /// Construct a JSON-serialisable representation of my appointment item.
        /// </summary>
        internal override ProtoItem<Outlook.AppointmentItem> CreateProtoItem(Outlook.AppointmentItem outlookItem)
        {
            return new ProtoAppointment(outlookItem);
        }

        public override void RemoveSynchronisationProperties()
        {
            olItem.ClearSynchronisationProperties();
        }

    }
}
