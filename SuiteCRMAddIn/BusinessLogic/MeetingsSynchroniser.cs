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
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public class MeetingsSynchroniser : AppointmentSyncing<MeetingSyncState>
    {
        public const string CrmModule = "Meetings";

        /// <summary>
        /// A cache of email addresses to CRM modules and identities
        /// </summary>
        private Dictionary<String, List<AddressResolutionData>> meetingRecipientsCache =
            new Dictionary<string, List<AddressResolutionData>>();


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
            catch (Exception any)
            {
                Log.Error("AppointmentSyncing.SetOutlookItemDuration", any);
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
                SetRecipients(olItem, crmItem, crmItem.GetValueAsString("id"), crmType);
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
            olItem.MeetingStatus = crmItem.GetValueAsString("assigned_user_id") == RestAPIWrapper.GetUserId() ?
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

            foreach (MeetingSyncState state in SyncStateManager.Instance.GetSynchronisedItems<MeetingSyncState>())
            {
                Outlook.AppointmentItem item = state.OutlookItem;

                try
                {
                    if (item.UserProperties[OrganiserPropertyName]?.Value == RestAPIWrapper.GetUserId() &&
                        item.Start > DateTime.Now)
                    {
                        result += AddOrUpdateMeetingAcceptanceFromOutlookToCRM(item);
                    }
                }
                catch (COMException comx)
                {
                    Log.Error($"Item with CRMid {state.CrmEntryId} appears to be invalid (HResult {comx.HResult})", comx);
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


        internal override string AddOrUpdateItemFromOutlookToCrm(SyncState<Outlook.AppointmentItem> syncState, string crmType, string entryId = "")
        {
            string result = base.AddOrUpdateItemFromOutlookToCrm(syncState, crmType, entryId);

            if (!string.IsNullOrEmpty(result))
            {
                if (string.IsNullOrEmpty(entryId))
                {
                    /* i.e. this was a new item saved to CRM for the first time */
                    SetCrmRelationshipFromOutlook(result, "Users", RestAPIWrapper.GetUserId());

                    this.SaveItem(syncState.OutlookItem);

                    if (syncState.OutlookItem.Recipients != null)
                    {
                        AddMeetingRecipientsFromOutlookToCrm(syncState.OutlookItem, result);
                    }

                    this.AddOrUpdateMeetingAcceptanceFromOutlookToCRM(syncState.OutlookItem);

                }
            }

            return result;
        }

        protected override MeetingSyncState AddNewItemFromCrmToOutlook(Outlook.MAPIFolder appointmentsFolder, string crmType, EntryValue crmItem, DateTime date_start)
        {
            var result = base.AddNewItemFromCrmToOutlook(appointmentsFolder, crmType, crmItem, date_start);

            SetRecipients(result.OutlookItem, crmItem, result.CrmEntryId, crmType);

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
            var meetingId = meeting.UserProperties[SyncStateManager.CrmIdPropertyName]?.Value;

            if (meetingId != null &&
                !string.IsNullOrEmpty(acceptance) &&
                SyncDirection.AllowOutbound(this.Direction))
            {
                foreach (AddressResolutionData resolution in this.ResolveRecipient(meeting, invitee))
                {
                    try
                    {
                        RestAPIWrapper.SetMeetingAcceptance(meetingId.ToString(), resolution.moduleName, resolution.moduleId, acceptance);
                        count++;
                    }
                    catch (Exception any)
                    {
                        this.Log.Error($"{this.GetType().Name}.AddOrUpdateMeetingAcceptanceFromOutlookToCRM: Failed to resolve invitee {smtpAddress}:", any);
                    }
                }
            }

            return count;
        }

        private void AddMeetingRecipientsFromOutlookToCrm(Outlook.AppointmentItem olItem, string meetingId)
        {
            LogItemAction(olItem, "AppointmentSyncing.AddMeetingRecipientsFromOutlookToCrm");
            foreach (Outlook.Recipient recipient in olItem.Recipients)
            {
                var smtpAddress = recipient.GetSmtpAddress();

                Log.Info($"recepientName= {recipient.Name}, recepient= {smtpAddress}");

                List<AddressResolutionData> resolutions = this.ResolveRecipient(olItem, recipient);

                foreach (AddressResolutionData resolution in resolutions)
                {
                    SetCrmRelationshipFromOutlook(meetingId, resolution);
                }
            }
        }



        protected void SetRecipients(Outlook.AppointmentItem olItem, EntryValue crmItem, string sMeetingID, string sModule)
        {
            this.LogItemAction(olItem, "SetRecipients");

            try
            {
                int iCount = olItem.Recipients.Count;
                for (int iItr = 1; iItr <= iCount; iItr++)
                {
                    olItem.Recipients.Remove(1);
                }

                foreach (var relationship in crmItem.relationships.link_list)
                {
                    foreach (LinkRecord record in relationship.records)
                    {
                        string email1 = record.data.GetValueAsString("email1");
                        string phone_work = record.data.GetValueAsString("phone_work");
                        string identifier = (sModule == MeetingsSynchroniser.CrmModule) || string.IsNullOrWhiteSpace(phone_work) ?
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
            finally
            {
                this.SaveItem(olItem);
            }
        }


        /// <summary>
        /// Sets up a CRM relationship to mimic an Outlook relationship
        /// </summary>
        /// <param name="meetingId">The ID of the meeting</param>
        /// <param name="recipient">The outlook recipient representing the person to link with.</param>
        /// <param name="foreignModule">the name of the module we're seeking to link with.</param>
        /// <returns>True if a relationship was created.</returns>
        protected string SetCrmRelationshipFromOutlook(string meetingId, Outlook.Recipient recipient, string foreignModule)
        {
            string foreignId = GetInviteeIdBySmtpAddress(recipient.GetSmtpAddress(), foreignModule);

            return !string.IsNullOrWhiteSpace(foreignId) &&
                SetCrmRelationshipFromOutlook(meetingId, foreignModule, foreignId) ?
                foreignId :
                string.Empty;
        }


        /// <summary>
        /// Sets up a CRM relationship to mimic an Outlook relationship
        /// </summary>
        /// <param name="meetingId">The meeting id.</param>
        /// <param name="resolution">Address resolution data from the cache.</param>
        /// <returns>True if a relationship was created.</returns>
        private bool SetCrmRelationshipFromOutlook(string meetingId, AddressResolutionData resolution)
        {
            return this.SetCrmRelationshipFromOutlook(meetingId, resolution.moduleName, resolution.moduleId);
        }


        /// <summary>
        /// Sets up a CRM relationship to mimic an Outlook relationship
        /// </summary>
        /// <param name="meetingId">The ID of the meeting</param>
        /// <param name="foreignModule">the name of the module we're seeking to link with.</param>
        /// <param name="foreignId">The id in the foreign module of the record we're linking to.</param>
        /// <returns>True if a relationship was created.</returns>
        private bool SetCrmRelationshipFromOutlook(string meetingId, string foreignModule, string foreignId)
        {
            bool result = false;

            if (foreignId != String.Empty)
            {
                SetRelationshipParams info = new SetRelationshipParams
                {
                    module2 = MeetingsSynchroniser.CrmModule,
                    module2_id = meetingId,
                    module1 = foreignModule,
                    module1_id = foreignId
                };
                result = RestAPIWrapper.SetRelationshipUnsafe(info);
            }

            return result;
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
                string meetingId = olItem.UserProperties[SyncStateManager.CrmIdPropertyName]?.Value;
                Dictionary<string, string> moduleIds = new Dictionary<string, string>();

                if (!string.IsNullOrEmpty(meetingId))
                {
                    foreach (string moduleName in new string[] { "Leads", "Users", ContactSynchroniser.CrmModule })
                    {
                        string moduleId = this.GetInviteeIdBySmtpAddress(smtpAddress, moduleName);
                        if (!string.IsNullOrEmpty(moduleId))
                        {
                            moduleIds[moduleName] = moduleId;
                            AddressResolutionData data = new AddressResolutionData(moduleName, moduleId, smtpAddress);
                            this.CacheAddressResolutionData(data);
                            result.Add(data);
                        }
                    }

                    if (moduleIds.ContainsKey(ContactSynchroniser.CrmModule))
                    {
                        string accountId = RestAPIWrapper.GetRelationship(
                            ContactSynchroniser.CrmModule,
                            moduleIds[ContactSynchroniser.CrmModule],
                            "accounts");

                        if (!string.IsNullOrWhiteSpace(accountId) &&
                            SetCrmRelationshipFromOutlook(meetingId, "Accounts", accountId))
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

        private bool TryAddRecipientInModule(string moduleName, string meetingId, Outlook.Recipient recipient)
        {
            bool result;
            string id = SetCrmRelationshipFromOutlook(meetingId, recipient, moduleName);

            if (!string.IsNullOrWhiteSpace(id))
            {
                string smtpAddress = recipient.GetSmtpAddress();

                this.CacheAddressResolutionData(
                    new AddressResolutionData(moduleName, id, smtpAddress));

                string accountId = RestAPIWrapper.GetRelationship(ContactSynchroniser.CrmModule, id, "accounts");

                if (!string.IsNullOrWhiteSpace(accountId) &&
                    SetCrmRelationshipFromOutlook(meetingId, "Accounts", accountId))
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


        /// <summary>
        /// Add an address resolution composed from this module name and record to the cache.
        /// </summary>
        /// <param name="moduleName">The name of the module in which the record was found</param>
        /// <param name="record">The record.</param>
        private void CacheAddressResolutionData(string moduleName, LinkRecord record)
        {
            Dictionary<string, object> data = record.data.AsDictionary();
            string smtpAddress = data[AddressResolutionData.EmailAddressFieldName].ToString();
            AddressResolutionData resolution = new AddressResolutionData(moduleName, data);

            CacheAddressResolutionData(resolution);
        }

        /// <summary>
        /// Add this resolution to the cache.
        /// </summary>
        /// <param name="resolution">The resolution to add.</param>
        private void CacheAddressResolutionData(AddressResolutionData resolution)
        {
            List<AddressResolutionData> resolutions;

            if (this.meetingRecipientsCache.ContainsKey(resolution.emailAddress))
            {
                resolutions = this.meetingRecipientsCache[resolution.emailAddress];
            }
            else
            {
                resolutions = new List<AddressResolutionData>();
                this.meetingRecipientsCache[resolution.emailAddress] = resolutions;
            }

            if (!resolutions.Any(x => x.moduleId == resolution.moduleId && x.moduleName == resolution.moduleName))
            {
                resolutions.Add(resolution);
            }

            Log.Debug($"Successfully cached recipient {resolution.emailAddress} => {resolution.moduleName}, {resolution.moduleId}.");
        }


        private void CacheAddressResolutionData(EntryValue crmItem)
        {
            foreach (var list in crmItem.relationships.link_list)
            {
                foreach (var record in list.records)
                {
                    var data = record.data.AsDictionary();
                    try
                    {
                        this.CacheAddressResolutionData(list.name, record);
                    }
                    catch (KeyNotFoundException kex)
                    {
                        Log.Error($"Key not found while caching meeting recipients.", kex);
                    }
                }
            }
        }


        /// <summary>
        /// Used for caching data for resolving email addresses to CRM records.
        /// </summary>
        private class AddressResolutionData
        {
            /// <summary>
            /// Expected name in the input map of the email address field.
            /// </summary>
            public const string EmailAddressFieldName = "email1";

            /// <summary>
            /// Expected name in the input map of the field containing the id in 
            /// the specified module.
            /// </summary>
            public const string ModuleIdFieldName = "id";

            /// <summary>
            /// Expected name in the input map of the field containing an associated id in 
            /// the `Accounts` module, if any.
            /// </summary>
            public const string AccountIdFieldName = "account_id";

            /// <summary>
            /// The email address resolved by this data.
            /// </summary>
            public readonly string emailAddress;
            /// <summary>
            /// The name of the CRM module to which it resolves.
            /// </summary>
            public readonly string moduleName;
            /// <summary>
            /// The id within that module of the record to which it resolves.
            /// </summary>
            public readonly string moduleId;

            /// <summary>
            /// The id within the `Accounts` module of a related record, if any.
            /// </summary>
            private readonly object accountId;

            public AddressResolutionData(string moduleName, string moduleId, string emailAddress)
            {
                this.moduleName = moduleName;
                this.moduleId = moduleId;
                this.emailAddress = emailAddress;
            }

            public AddressResolutionData(string moduleName, Dictionary<string, object> data)
            {
                this.moduleName = moduleName;
                this.moduleId = data[ModuleIdFieldName]?.ToString();
                this.emailAddress = data[EmailAddressFieldName]?.ToString();
                try
                {
                    this.accountId = data[AccountIdFieldName]?.ToString();
                }
                catch (KeyNotFoundException)
                {
                    // and ignore it; that key often won't be there.
                }
            }
        }
    }
}
