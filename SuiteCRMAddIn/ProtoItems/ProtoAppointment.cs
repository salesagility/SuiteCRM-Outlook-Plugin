
namespace SuiteCRMAddIn.ProtoItems
{
    using BusinessLogic;
    using SuiteCRMClient;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Broadly, a C# representation of a CRM appointment.
    /// </summary>
    public class ProtoAppointment : ProtoItem<Outlook.AppointmentItem>
    {
        private string Body;
        private int Duration;
        private DateTime End;
        private string Location;
        private string Organiser;
        private DateTime Start;
        private string Subject;

        public ProtoAppointment(Outlook.AppointmentItem olItem)
        {
            this.Body = olItem.Body;
            this.Duration = olItem.Duration;
            this.End = olItem.End;
            this.Location = olItem.Location;
            this.Start = olItem.Start;
            this.Subject = olItem.Subject;

            var organiserProperty = olItem.UserProperties[AppointmentSyncing.OrganiserPropertyName];

            if (organiserProperty == null || string.IsNullOrWhiteSpace(organiserProperty.Value))
            {
                 this.Organiser = clsSuiteCRMHelper.GetUserId();
            }
            else
            {
                this.Organiser = organiserProperty.Value.ToString();
            }
        }

        public override NameValueCollection AsNameValues(string entryId)
        {
            NameValueCollection data = new NameValueCollection();

            data.Add(clsSuiteCRMHelper.SetNameValuePair("name", this.Subject));
            data.Add(clsSuiteCRMHelper.SetNameValuePair("description", this.Body));
            data.Add(clsSuiteCRMHelper.SetNameValuePair("location", this.Location));
            data.Add(clsSuiteCRMHelper.SetNameValuePair("date_start", string.Format("{0:yyyy-MM-dd HH:mm:ss}", this.Start.ToUniversalTime())));
            data.Add(clsSuiteCRMHelper.SetNameValuePair("date_end", string.Format("{0:yyyy-MM-dd HH:mm:ss}", this.End.ToUniversalTime())));
            data.Add(clsSuiteCRMHelper.SetNameValuePair("duration_minutes", (this.Duration % 60).ToString()));
            data.Add(clsSuiteCRMHelper.SetNameValuePair("duration_hours", (this.Duration / 60).ToString()));
            data.Add(clsSuiteCRMHelper.SetNameValuePair("assigned_user_id", this.Organiser));

            if (!string.IsNullOrEmpty(entryId))
            {
                data.Add(clsSuiteCRMHelper.SetNameValuePair("id", entryId));
            }

            return data;
        }
    }
}
