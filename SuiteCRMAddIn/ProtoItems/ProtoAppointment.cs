
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
        /// <summary>
        /// Header for a block of accept/decline links in a meeting invite body.
        /// </summary>
        public const string AcceptDeclineHeader = "-- \nAccept/Decline links";

        private string Body;
        private int Duration;
        private DateTime End;
        private string Location;
        private string Organiser;
        private DateTime Start;
        private string Subject;

        public ProtoAppointment(Outlook.AppointmentItem olItem)
        {
            this.Body = TextUtilities.StripAndTruncate(olItem.Body, AcceptDeclineHeader);
            this.Duration = olItem.Duration;
            this.End = olItem.End;
            this.Location = olItem.Location;
            this.Start = olItem.Start;
            this.Subject = olItem.Subject;

            var organiserProperty = olItem.UserProperties[AppointmentSyncing.OrganiserPropertyName];

            if (organiserProperty == null || string.IsNullOrWhiteSpace(organiserProperty.Value))
            {
                 this.Organiser = RestAPIWrapper.GetUserId();
            }
            else
            {
                this.Organiser = organiserProperty.Value.ToString();
            }
        }

        public override NameValueCollection AsNameValues(string entryId)
        {
            NameValueCollection data = new NameValueCollection();

            data.Add(RestAPIWrapper.SetNameValuePair("name", this.Subject));
            data.Add(RestAPIWrapper.SetNameValuePair("description", this.Body));
            data.Add(RestAPIWrapper.SetNameValuePair("location", this.Location));
            data.Add(RestAPIWrapper.SetNameValuePair("date_start", string.Format("{0:yyyy-MM-dd HH:mm:ss}", this.Start.ToUniversalTime())));
            data.Add(RestAPIWrapper.SetNameValuePair("date_end", string.Format("{0:yyyy-MM-dd HH:mm:ss}", this.End.ToUniversalTime())));
            data.Add(RestAPIWrapper.SetNameValuePair("duration_minutes", (this.Duration % 60).ToString()));
            data.Add(RestAPIWrapper.SetNameValuePair("duration_hours", (this.Duration / 60).ToString()));
            data.Add(RestAPIWrapper.SetNameValuePair("assigned_user_id", this.Organiser));

            if (!string.IsNullOrEmpty(entryId))
            {
                data.Add(RestAPIWrapper.SetNameValuePair("id", entryId));
            }

            return data;
        }
    }
}
