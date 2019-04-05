namespace SuiteCRMAddIn.ProtoItems
{
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// A C# proxy for a CRM object.
    /// </summary>
    /// <remarks>
    /// ProtoItems have two jobs: one is acting as a cache for values read from a 
    /// CRM item, and the other is to act as a transducer between CRM items and 
    /// Outlook items.
    /// </remarks>
    /// <typeparam name="OutlookItemType">The type of outlook item this proto-item transduces.</typeparam>
    public abstract class ProtoItem<OutlookItemType> : ProtoItem
    {
        /// <summary>
        /// (For meetings only) the meeting status.
        /// </summary>
        public Outlook.OlMeetingStatus Status;

        /// <summary>
        /// Super-Constructor for a ProtoItem which is not a meeting; sets status to `olNonMeeting`.
        /// </summary>
        public ProtoItem() : this(Outlook.OlMeetingStatus.olNonMeeting) { }

        public ProtoItem(Outlook.OlMeetingStatus meetingStatus)
        {
            this.Status = meetingStatus;
        }
    }
}