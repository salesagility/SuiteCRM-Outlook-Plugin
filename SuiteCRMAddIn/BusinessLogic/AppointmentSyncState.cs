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
    using Exceptions;

    public abstract class AppointmentSyncState: SyncState<Outlook.AppointmentItem>
    {
        public AppointmentSyncState(Outlook.AppointmentItem item, CrmId crmId, DateTime modifiedDate) : base(item, item.EntryID, crmId, modifiedDate)
        {
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
                if (CrmId.IsInvalid(this.CrmEntryId)) { this.CrmEntryId = CrmId.Empty; }

                StringBuilder bob = new StringBuilder();
                bob.Append($"\tOutlook Id  : {OutlookItem.EntryID}\n\tCRM Id      : {this.CrmEntryId}\n\tSubject     : '{OutlookItem.Subject}'\n\tSensitivity : {OutlookItem.Sensitivity}\n\tStatus     : {OutlookItem.MeetingStatus}\n\tReminder set {OutlookItem.ReminderSet}\n\tState      : {this.TxState}\n\tRecipients:\n");
                foreach (Outlook.Recipient recipient in OutlookItem.Recipients)
                {
                    bob.Append($"\t\t{recipient.Name}: {recipient.GetSmtpAddress()} - ({recipient.MeetingResponseStatus})\n");
                }

                return bob.ToString();
            }
        }

        public override Outlook.OlSensitivity OutlookItemSensitivity => 
            OutlookItem != null && OutlookItem.IsValid() ? OutlookItem.Sensitivity : Outlook.OlSensitivity.olPrivate;

        public override Outlook.UserProperties OutlookUserProperties => 
            OutlookItem != null && OutlookItem.IsValid() ? OutlookItem.UserProperties : null;

        /// <summary>
        /// #6034: occasionally we get spurious ItemChange events where the 
        /// value of Duration appear as zero, although nothing has occured to
        /// make this change. This is a hack around the problem while we try
        /// to understand it better.
        /// </summary>
        /// <returns>false if duration was set to zero; as a side effect,
        /// resets Duration to its last known good value.</returns>
        internal override bool ShouldPerformSyncNow()
        {
            bool result;

            try
            {
                result = base.ShouldPerformSyncNow();
            }
            catch (DurationSetToZeroException dz)
            {
                Outlook.AppointmentItem appt = (this.OutlookItem as Outlook.AppointmentItem);

                if (appt != null)
                {
                    appt.Duration = dz.Duration;
                }

                result = false;
            }

            return result;
        }

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
            if (this.OutlookItem != null && this.OutlookItem.IsValid()) this.OutlookItem?.Save();
        }

        public override bool VerifyItem()
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
