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
    using SuiteCRMClient;

    public abstract class AppointmentSyncState: SyncState<Outlook.AppointmentItem>
    {
        public AppointmentSyncState(Outlook.AppointmentItem item, CrmId crmId, DateTime modifiedDate) : base(item, crmId, modifiedDate)
        {
            this.outlookItemId = item.EntryID;
        }

        /// <summary>
        /// When we're asked for the CrmType the underlying object may have ceased to
        /// exist - so cache it!
        /// </summary>
        private string crmType;


        public override Outlook.Folder DefaultFolder => (Outlook.Folder)MapiNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);


        /// <summary>
        /// The CRM type of the item I represent.
        /// </summary>
        public override string CrmType
        {
            get
            {
                try
                {
                    switch (OutlookItem.MeetingStatus)
                    {
                        case Outlook.OlMeetingStatus.olNonMeeting:
                            crmType = CallsSynchroniser.CrmModule;
                            break;
                        default:
                            crmType = MeetingsSynchroniser.CrmModule;
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
                CrmId crmId = OutlookItem.GetCrmId();
                if (CrmId.IsInvalid(crmId)) { crmId = CrmId.Empty; }

                StringBuilder bob = new StringBuilder();
                bob.Append($"\tOutlook Id  : {OutlookItem.EntryID}\n\tCRM Id      : {crmId}\n\tSubject     : '{OutlookItem.Subject}'\n\tSensitivity : {OutlookItem.Sensitivity}\n\tStatus     : {OutlookItem.MeetingStatus}\n\tReminder set {OutlookItem.ReminderSet}\n\tState      : {this.TxState}\n\tRecipients:\n");
                foreach (Outlook.Recipient recipient in OutlookItem.Recipients)
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

        public override void RemoveSynchronisationProperties()
        {
            OutlookItem.ClearSynchronisationProperties();
        }

        internal override void SaveItem()
        {
            this.OutlookItem?.Save();
        }

        protected override bool VerifyItem()
        {
            bool result;
            try
            {
                result = !string.IsNullOrEmpty(this.Item?.EntryID);
            }
            catch (Exception ex) when (ex is InvalidComObjectException || ex is COMException)
            {
                result = false;
            }

            return result;
        }
    }
}
