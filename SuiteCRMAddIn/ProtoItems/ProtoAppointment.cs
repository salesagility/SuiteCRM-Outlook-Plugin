
namespace SuiteCRMAddIn.ProtoItems
{
    using BusinessLogic;
    using SuiteCRMClient;
    using SuiteCRMClient.RESTObjects;
    using SuiteCRMClient.Logging;
    using System;
    using System.Linq;
    using System.Net.Mail;
    using System.Text.RegularExpressions;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System.Collections.Generic;
    using SuiteCRMAddIn.Extensions;

    /// <summary>
    /// Broadly, a C# representation of a CRM appointment.
    /// </summary>
    public class ProtoAppointment<SyncStateType> : ProtoItem<Outlook.AppointmentItem>
        where SyncStateType : SyncState<Outlook.AppointmentItem>
    {
        private readonly string body;
        private readonly int duration;
        private readonly DateTime end;
        private readonly string location;
        private readonly CrmId organiser;
        private readonly DateTime start;
        private readonly string subject;
        private readonly string globalId;
        private readonly string CancelledPrefix = "CANCELLED";
        private readonly ISet<string> recipientAddresses = new HashSet<string>();
        private CrmId CrmEntryId;

        private readonly Outlook.AppointmentItem olItem;

        /// <summary>
        /// Public read access for duration.
        /// </summary>
        public int Duration => this.duration;

        /// <summary>
        /// Readonly access to an ordered list of my recipient addresses.
        /// </summary>
        public List<string> RecipientAddresses
        {
            get
            {
                return recipientAddresses.AsEnumerable().OrderBy(x => x).ToList();
            }
        }

        public override string Description
        {
            get
            {
                return $"{subject} ({start})";
            }
        }


        /// <summary>
        /// Create a new instance of ProtoAppointment, taking values from this Outlook item.
        /// </summary>
        /// <param name="olItem">The Outlook item to take values from.</param>
        public ProtoAppointment(Outlook.AppointmentItem olItem) : base(olItem.MeetingStatus)
        {
            this.olItem = olItem;
            this.body = olItem.Body;
            this.CrmEntryId = olItem.GetCrmId();
            this.duration = olItem.Duration;
            this.end = olItem.End;
            this.location = olItem.Location;
            this.start = olItem.Start;
            this.subject = olItem.Subject;
            this.globalId = olItem.GlobalAppointmentID;
 
            var organiserProperty = olItem.UserProperties[AppointmentsSynchroniser<SyncStateType>.OrganiserPropertyName];

            if (organiserProperty == null || string.IsNullOrWhiteSpace(organiserProperty.Value))
            {
                if (olItem.Organizer == Globals.ThisAddIn.Application.GetCurrentUsername())
                {
                    this.organiser = CrmId.Get(RestAPIWrapper.GetUserId());
                }
                else
                {
                    this.organiser = TryResolveOrganiser(olItem);
                }
            }
            else
            {
                this.organiser = CrmId.Get(organiserProperty.Value.ToString());
            }

            foreach (Outlook.Recipient recipient in olItem.Recipients)
            {
                this.recipientAddresses.Add(recipient.GetSmtpAddress());
            }
        }

        /// <summary>
        /// Try to resolve the organiser of this Outlook Item against the users of the CRM.
        /// </summary>
        /// <param name="olItem">The Outlook item representing a meeting.</param>
        /// <returns>The id of the related user if any, else the empty string.</returns>
        public static CrmId TryResolveOrganiser(Outlook.AppointmentItem olItem)
        {
            CrmId result = CrmId.Empty;
            string organiser = olItem.Organizer;

            try
            {
                if (organiser.IndexOf('@') > -1)
                {
                    foreach (string pattern in new string[] { @".*<(.+@.+)>", @".+@.+" })
                    {
                        Match match = Regex.Match(organiser, pattern, RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            string address = match.Groups[0].Value;

                            try
                            {
                                result = CrmId.Get(RestAPIWrapper.GetUserId(new MailAddress(address)));
                            }
                            catch (FormatException)
                            {
                                // not a valid email address - no matter.
                            }
                        }
                    }
                }
                else
                {
                    result = CrmId.Get(RestAPIWrapper.GetUserId(organiser));
                }
            }
            catch (Exception any)
            {
                ErrorHandler.Handle($"Failed to resolve organiser `{olItem.Organizer}` of meeting `{olItem.Subject}`", any);
            }

            return result;
        }


        /// <summary>
        /// AsNameValues is used in transmission to CRM as well as for comparison, so it should NOT
        /// access our cache of recipient addresses.
        /// </summary>
        /// <returns>A set of name/value pairs suitable for transmitting to CRM.</returns>
        public override NameValueCollection AsNameValues()
        {
            string statusString;
            string name;

            switch (this.Status)
            {
                case Outlook.OlMeetingStatus.olMeetingCanceled:
                    statusString = "Not Held";
                    name = this.subject.StartsWith(CancelledPrefix) ? this.subject : $"{CancelledPrefix}: {this.subject}";
                    break;
                default:
                    statusString = this.start < DateTime.Now ? "Held" : "Planned";
                    name = this.subject;
                    break;
            }

            NameValueCollection data = new NameValueCollection
            {
                RestAPIWrapper.SetNameValuePair("name", name),
                RestAPIWrapper.SetNameValuePair("description", this.body),
                RestAPIWrapper.SetNameValuePair("location", this.location),
                RestAPIWrapper.SetNameValuePair("date_start",
                    string.Format("{0:yyyy-MM-dd HH:mm:ss}", this.start.ToUniversalTime())),
                RestAPIWrapper.SetNameValuePair("date_end",
                    string.Format("{0:yyyy-MM-dd HH:mm:ss}", this.end.ToUniversalTime())),
                RestAPIWrapper.SetNameValuePair("duration_minutes", (this.duration % 60).ToString()),
                RestAPIWrapper.SetNameValuePair("duration_hours", (this.duration / 60).ToString()),
                RestAPIWrapper.SetNameValuePair("outlook_id", this.globalId),
                RestAPIWrapper.SetNameValuePair("status", statusString)
            };

            if (CrmId.IsValid(this.organiser))
            {
                data.Add(RestAPIWrapper.SetNameValuePair("assigned_user_id", this.organiser.ToString()));
            }

            if (CrmId.IsInvalid(CrmEntryId))
            {
                /* A Guid can be constructed from a 32 digit hex string. The globalId is a
                 * 112 digit hex string. It appears from inspection that the least significant
                 * bytes are those that vary between examples, with the most significant bytes 
                 * being invariant in the samples we have to hand. */
                CrmEntryId = CrmId.Get(new Guid(this.globalId.Substring(this.globalId.Length - 32)));
                data.Add(RestAPIWrapper.SetNameValuePair("new_with_id", true));
            }

            data.Add(RestAPIWrapper.SetNameValuePair("id", CrmEntryId.ToString()));

            return data;
        }
    }
}
