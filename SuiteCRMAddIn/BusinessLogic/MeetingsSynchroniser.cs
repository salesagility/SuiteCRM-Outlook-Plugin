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
    using Microsoft.Office.Interop.Outlook;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public class MeetingsSynchroniser : AppointmentsSynchroniser<MeetingSyncState>
    {
        public const string CrmModule = "Meetings";

        public MeetingsSynchroniser(string name, SyncContext context) : base(name, context)
        {
        }

        public override string DefaultCrmModule
        {
            get
            {
                return CrmModule;
            }
        }

        public override SyncDirection.Direction Direction => Properties.Settings.Default.SyncMeetings;

        /// <summary>
        /// Specialisation: also fetch names and email ids of recipients.
        /// </summary>
        /// <param name="offset">The offset into the resultset at which the page begins.</param>
        /// <returns>A set of entries.</returns>
        protected override EntryList GetEntriesPage(int offset)
        {
            return RestAPIWrapper.GetEntryList(this.DefaultCrmModule,
                String.Format(fetchQueryPrefix, RestAPIWrapper.GetUserId()),
                Properties.Settings.Default.SyncMaxRecords, "date_start DESC", offset, false,
                RestAPIWrapper.GetSugarFields(this.DefaultCrmModule), new[] {
                    new { @name = "users", @value = new[] {"id", "email1", "phone_work" } },
                    new { @name = "contacts", @value = new[] {"id", "account_id", "email1", "phone_work" } },
                    new { @name = "leads", @value = new[] {"id", "email1", "phone_work" } }
                    });
        }

        /// <summary>
        /// Specialisation: also set end time and location.
        /// </summary>
        /// <param name="crmItem">The CRM item.</param>
        /// <param name="olItem">The Outlook item.</param>
        protected override void SetOutlookItemDuration(EntryValue crmItem, Outlook.AppointmentItem olItem)
        {
            try
            {
                base.SetOutlookItemDuration(crmItem, olItem);
                olItem.Location = crmItem.GetValueAsString("location");
                olItem.End = olItem.Start.AddMinutes(olItem.Duration);
            }
            catch (System.Exception any)
            {
                ErrorHandler.Handle("Failed while setting Outlook item duration", any);
            }
        }

       protected override void UpdateOutlookDetails(string crmType, EntryValue crmItem, DateTime date_start, Outlook.AppointmentItem olItem)
        {
            try
            {
                olItem.Start = date_start;
                var minutesString = crmItem.GetValueAsString("duration_minutes");
                var hoursString = crmItem.GetValueAsString("duration_hours");

                int minutes = string.IsNullOrWhiteSpace(minutesString) ? 0 : int.Parse(minutesString);
                int hours = string.IsNullOrWhiteSpace(hoursString) ? 0 : int.Parse(hoursString);

                olItem.Duration = minutes + hours * 60;

                olItem.Location = crmItem.GetValueAsString("location");
                olItem.End = olItem.Start;
                if (hours > 0)
                {
                    olItem.End.AddHours(hours);
                }
                if (minutes > 0)
                {
                    olItem.End.AddMinutes(minutes);
                }
                SetOutlookRecipientsFromCRM(olItem, crmItem, crmItem.GetValueAsString("id"), crmType);
            }
            finally
            {
                this.SaveItem(olItem);
            }
        }

        protected override bool ShouldAddOrUpdateItemFromCrmToOutlook(Outlook.MAPIFolder folder, string crmType, EntryValue crmItem)
        {
            return crmType == "Meetings";
        }

        protected override void SetMeetingStatus(Outlook.AppointmentItem olItem, EntryValue crmItem)
        {
            olItem.MeetingStatus = CrmId.Get(crmItem.GetValueAsString("assigned_user_id")).Equals( RestAPIWrapper.GetUserId()) ?
                Outlook.OlMeetingStatus.olMeeting :
                Outlook.OlMeetingStatus.olMeetingReceived;
        }

        protected override void OtherIterationActions()
        {
            CheckMeetingAcceptances();
        }

        /// <summary>
        /// Check meeting acceptances for all future meetings.
        /// </summary>
        private int CheckMeetingAcceptances()
        {
            int result = 0;

            foreach (MeetingSyncState state in SyncStateManager.Instance.GetSynchronisedItems<MeetingSyncState>().Where(s => s.VerifyItem()))
            {
                Outlook.AppointmentItem item = state.OutlookItem;

                try
                {
                    if (CrmId.Get(item.UserProperties[OrganiserPropertyName]?.Value)
                            .Equals(RestAPIWrapper.GetUserId()) &&
                        item.Start > DateTime.Now)
                    {
                        result += AddOrUpdateMeetingAcceptanceFromOutlookToCRM(item);
                    }
                }
                catch (TypeInitializationException tix)
                {
                    Log.Warn("Failed to create CrmId with value '{item.UserProperties[OrganiserPropertyName]?.Value}'", tix);
                }
                catch (COMException comx)
                {
                    ErrorHandler.Handle($"Item with CRMid {state.CrmEntryId} appears to be invalid (HResult {comx.HResult})", comx);
                    this.HandleItemMissingFromOutlook(state);
                }
            }

            return result;
        }


        /// <summary>
        /// Check meeting acceptances for the invitees of this `meeting`.
        /// </summary>
        /// <param name="meeting">The meeting.</param>
        /// <returns>the number of valid acceptance statuses found.</returns>
        public int UpdateMeetingAcceptances(Outlook.MeetingItem meeting)
        {
            return meeting == null ?
                0 :
                this.AddOrUpdateMeetingAcceptanceFromOutlookToCRM(meeting.GetAssociatedAppointment(false));
        }

        /// <summary>
        ///     Override: we get notified of a removal, for a Meeting item, when the meeting is
        ///     cancelled. We do NOT want to remove such an item; instead, we want to update it.
        /// </summary>
        /// <param name="state"></param>
        protected override void RemoveFromCrm(SyncState state)
        {
            var meeting = state as MeetingSyncState;

            if (meeting != null)
            {
                meeting.Cache.Status = OlMeetingStatus.olMeetingCanceled;

                RestAPIWrapper.SetEntry(meeting.Cache.AsNameValues(), DefaultCrmModule);
            }
        }

        internal override CrmId AddOrUpdateItemFromOutlookToCrm(SyncState<Outlook.AppointmentItem> syncState)
        {
            CrmId previousCrmId = syncState.CrmEntryId;
            CrmId result = base.AddOrUpdateItemFromOutlookToCrm(syncState);

            if (CrmId.IsValid(result))
            {
                if (CrmId.IsInvalid(previousCrmId)) /* i.e., it's new */
                {
                    if (syncState.OutlookItem.Recipients != null)
                    {
                        AddMeetingRecipientsFromOutlookToCrm(syncState.OutlookItem, result);
                    }

                    this.AddOrUpdateMeetingAcceptanceFromOutlookToCRM(syncState.OutlookItem);

                }
            }

            return result;
        }


        /// <summary>
        /// Set the meeting acceptance status, in CRM, of all invitees to this meeting from
        /// their acceptance status in Outlook.
        /// </summary>
        private int AddOrUpdateMeetingAcceptanceFromOutlookToCRM(Outlook.AppointmentItem meeting)
        {
            int count = 0;
            foreach (Outlook.Recipient recipient in meeting.Recipients)
            {
                var acceptance = recipient.CrmAcceptanceStatus();
                if (!string.IsNullOrEmpty(acceptance))
                {
                    count += this.AddOrUpdateMeetingAcceptanceFromOutlookToCRM(meeting, recipient, acceptance);
                }
            }

            return count;
        }

        /// <summary>
        /// Set the meeting acceptance status, in CRM, for this invitee to this meeting from
        /// their acceptance status in Outlook.
        /// </summary>
        /// <param name="meeting">The appointment item representing the meeting</param>
        /// <param name="invitee">The recipient item representing the invitee</param>
        /// <param name="acceptance">The acceptance status of this invitee of this meeting 
        /// as a string recognised by CRM.</param>
        /// 
        private int AddOrUpdateMeetingAcceptanceFromOutlookToCRM(Outlook.AppointmentItem meeting, Outlook.Recipient invitee, string acceptance)
        {
            int count = 0;
            string smtpAddress = invitee.GetSmtpAddress();
            var meetingId = meeting.GetCrmId();

            if (meetingId != null &&
                !string.IsNullOrEmpty(acceptance) &&
                SyncDirection.AllowOutbound(this.Direction))
            {
                foreach (AddressResolutionData resolution in this.ResolveRecipient(meeting, invitee))
                {
                    try
                    {
                        RestAPIWrapper.SetMeetingAcceptance(meetingId.ToString(), resolution.ModuleName, resolution.ModuleId.ToString(), acceptance);
                        count++;
                    }
                    catch (System.Exception any)
                    {
                        ErrorHandler.Handle($"Failed to resolve meeting invitee {smtpAddress}:", any);
                    }
                }
            }

            return count;
        }

        private void AddMeetingRecipientsFromOutlookToCrm(Outlook.AppointmentItem olItem, CrmId meetingId)
        {
            LogItemAction(olItem, "AppointmentSyncing.AddMeetingRecipientsFromOutlookToCrm");
            foreach (Outlook.Recipient recipient in olItem.Recipients)
            {
                var smtpAddress = recipient.GetSmtpAddress();

                Log.Info($"recepientName= {recipient.Name}, recepient= {smtpAddress}");

                List<AddressResolutionData> resolutions = this.ResolveRecipient(olItem, recipient);

                foreach (AddressResolutionData resolution in resolutions)
                {
                    SetCrmRelationshipFromOutlook(this, meetingId, resolution);
                }
            }
        }


        /// <summary>
        /// Set up the recipients of the appointment represented by this `olItem` from this `crmItem`.
        /// </summary>
        /// <param name="olItem">The Outlook item to update.</param>
        /// <param name="crmItem">The CRM item to use as source.</param>
        /// <param name="sMeetingID"></param>
        /// <param name="crmModule">The module the CRM item is in.</param>
        protected void SetOutlookRecipientsFromCRM(Outlook.AppointmentItem olItem, EntryValue crmItem, string sMeetingID, string crmModule)
        {
            this.LogItemAction(olItem, "SetRecipients");

            try
            {
                int iCount = olItem.Recipients.Count;
                for (int iItr = 1; iItr <= iCount; iItr++)
                {
                    olItem.Recipients.Remove(1);
                }

                if (crmItem != null && crmItem.relationships != null && crmItem.relationships.link_list != null)
                {
                    foreach (var relationship in crmItem.relationships.link_list)
                    {
                        foreach (LinkRecord record in relationship.records)
                        {
                            string email1 = record.data.GetValueAsString("email1");
                            string phone_work = record.data.GetValueAsString("phone_work");
                            string identifier = (crmModule == MeetingsSynchroniser.CrmModule) || string.IsNullOrWhiteSpace(phone_work) ?
                                    email1 :
                                    $"{email1} : {phone_work}";

                            if (!String.IsNullOrWhiteSpace(identifier))
                            {
                                if (olItem.GetOrganizer().GetSmtpAddress() != email1)
                                {
                                    olItem.EnsureRecipient(email1, identifier);
                                }
                            }
                        }
                    }
                }
            }
            finally
            {
                this.SaveItem(olItem);
            }
        }



        /// <summary>
        /// Find all CRM records related to this recipient of this meeting, and produce address
        /// resolution data from them.
        /// </summary>
        /// <param name="olItem">An appointment, assumed to be a meeting.</param>
        /// <param name="recipient">A recipient of that meeting request.</param>
        /// <returns>A list of address resolution objects.</returns>
        private List<AddressResolutionData> ResolveRecipient(Outlook.AppointmentItem olItem, Outlook.Recipient recipient)
        {
            List<AddressResolutionData> result = new List<AddressResolutionData>();
            var smtpAddress = recipient.GetSmtpAddress();

            Log.Info($"recepientName= {recipient.Name}, recepient= {smtpAddress}");

            if (this.meetingRecipientsCache.ContainsKey(smtpAddress))
            {
                result.AddRange(meetingRecipientsCache[smtpAddress]);
            }
            else
            {
                CrmId meetingId = olItem.GetCrmId();
                Dictionary<string, CrmId> moduleIds = new Dictionary<string, CrmId>();

                if (CrmId.IsValid(meetingId))
                {
                    foreach (string moduleName in new string[] { "Leads", "Users", ContactSynchroniser.CrmModule })
                    {
                        CrmId moduleId = this.GetInviteeIdBySmtpAddress(smtpAddress, moduleName);
                        if (CrmId.IsValid(moduleId))
                        {
                            moduleIds[moduleName] = moduleId;
                            AddressResolutionData data = new AddressResolutionData(moduleName, moduleId, smtpAddress);
                            this.CacheAddressResolutionData(data);
                            result.Add(data);
                        }
                    }

                    if (moduleIds.ContainsKey(ContactSynchroniser.CrmModule))
                    {
                        CrmId accountId = CrmId.Get(RestAPIWrapper.GetRelationship(
                            ContactSynchroniser.CrmModule,
                            moduleIds[ContactSynchroniser.CrmModule].ToString(),
                            "accounts"));

                        if (CrmId.IsValid(accountId) &&
                            SetCrmRelationshipFromOutlook(this, meetingId, "Accounts", accountId))
                        {
                            var data = new AddressResolutionData("Accounts", accountId, smtpAddress);
                            this.CacheAddressResolutionData(data);
                            result.Add(data);
                        }
                    }
                }
            }

            return result;
        }

        private bool TryAddRecipientInModule(string moduleName, CrmId meetingId, Outlook.Recipient recipient)
        {
            bool result;
            CrmId id = SetCrmRelationshipFromOutlook(this, meetingId, recipient, moduleName);

            if (CrmId.IsValid(id))
            {
                string smtpAddress = recipient.GetSmtpAddress();

                this.CacheAddressResolutionData(
                    new AddressResolutionData(moduleName, id, smtpAddress));

                CrmId accountId = CrmId.Get(RestAPIWrapper.GetRelationship(
                    ContactSynchroniser.CrmModule, id.ToString(), "accounts"));

                if (CrmId.IsValid(accountId) &&
                    SetCrmRelationshipFromOutlook(this, meetingId, "Accounts", accountId))
                {
                    this.CacheAddressResolutionData(
                        new AddressResolutionData("Accounts", accountId, smtpAddress));
                }

                result = true;
            }
            else
            {
                result = false;
            }

            return result;
        }

        protected override SyncState<Outlook.AppointmentItem> AddOrUpdateItemFromCrmToOutlook(Outlook.MAPIFolder folder, string crmType, EntryValue crmItem)
        {
            var result = base.AddOrUpdateItemFromCrmToOutlook(folder, crmType, crmItem);

            if (crmItem?.relationships?.link_list != null)
            {
                CacheAddressResolutionData(crmItem);
            }

            return result;
        }
    }
}
