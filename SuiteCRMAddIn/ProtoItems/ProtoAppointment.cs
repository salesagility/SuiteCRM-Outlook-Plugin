
namespace SuiteCRMAddIn.ProtoItems
{
    using BusinessLogic;
    using SuiteCRMClient;
    using SuiteCRMClient.RESTObjects;
    using System;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Broadly, a C# representation of a CRM appointment.
    /// </summary>
    public class ProtoAppointment : ProtoItem<Outlook.AppointmentItem>
    {
        private readonly string body;
        private readonly int duration;
        private readonly DateTime end;
        private readonly string location;
        private readonly string organiser;
        private readonly DateTime start;
        private readonly string subject;
        private readonly string outlookId;
        private readonly Outlook.OlMeetingStatus status;
        private readonly string CancelledPrefix = "CANCELLED";

        public ProtoAppointment(Outlook.AppointmentItem olItem)
        {
            this.body = olItem.Body;
            this.duration = olItem.Duration;
            this.end = olItem.End;
            this.location = olItem.Location;
            this.start = olItem.Start;
            this.subject = olItem.Subject;
            this.outlookId = olItem.EntryID;
            this.status = olItem.MeetingStatus;

            var organiserProperty = olItem.UserProperties[AppointmentSyncing.OrganiserPropertyName];

            if (organiserProperty == null || string.IsNullOrWhiteSpace(organiserProperty.Value))
            {
                 this.organiser = RestAPIWrapper.GetUserId();
            }
            else
            {
                this.organiser = organiserProperty.Value.ToString();
            }
        }

        public override NameValueCollection AsNameValues(string entryId)
        {
            NameValueCollection data = new NameValueCollection();
            string statusString;
            string name;

            switch (this.status)
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

            data.Add(RestAPIWrapper.SetNameValuePair("name", name));
            data.Add(RestAPIWrapper.SetNameValuePair("description", this.body));
            data.Add(RestAPIWrapper.SetNameValuePair("location", this.location));
            data.Add(RestAPIWrapper.SetNameValuePair("date_start", string.Format("{0:yyyy-MM-dd HH:mm:ss}", this.start.ToUniversalTime())));
            data.Add(RestAPIWrapper.SetNameValuePair("date_end", string.Format("{0:yyyy-MM-dd HH:mm:ss}", this.end.ToUniversalTime())));
            data.Add(RestAPIWrapper.SetNameValuePair("duration_minutes", (this.duration % 60).ToString()));
            data.Add(RestAPIWrapper.SetNameValuePair("duration_hours", (this.duration / 60).ToString()));
            data.Add(RestAPIWrapper.SetNameValuePair("assigned_user_id", this.organiser));
            data.Add(RestAPIWrapper.SetNameValuePair("outlook_id", this.outlookId));
            data.Add(RestAPIWrapper.SetNameValuePair("status", statusString));

            if (!string.IsNullOrEmpty(entryId))
            {
                data.Add(RestAPIWrapper.SetNameValuePair("id", entryId));
            }

            return data;
        }
    }
}
