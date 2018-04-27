namespace SuiteCRMAddIn.Daemon
{
    using Exceptions;
    using SuiteCRMAddIn.BusinessLogic;
    using System;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public class UpdateMeetingAcceptancesAction : AbstractDaemonAction
    {
        private readonly MeetingsSynchroniser synchroniser;
        private readonly Outlook.MeetingItem meeting;

        public UpdateMeetingAcceptancesAction(
            MeetingsSynchroniser synchroniser,
            Outlook.MeetingItem meeting) : base(5)
        {
            this.synchroniser = synchroniser;
            this.meeting = meeting;
        }

        public override string Description
        {
            get
            {
                return $"Checking acceptances for meeting `{this.meeting.Subject}`";
            }
        }

        public override string Perform()
        {
            throw new ActionDisabledException();
            //if (this.synchroniser.UpdateMeetingAcceptances(this.meeting) == 0)
            //{
            //    throw new ActionRetryableException($"Meeting `{this.meeting.Subject}`: no acceptances yet");
            //}
            //return ("OK");
        }
    }
}
