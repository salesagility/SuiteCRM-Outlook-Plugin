namespace SuiteCRMAddIn.BusinessLogic
{
    /// <summary>
    /// States a SyncState object can be in with regard to transmission and synchronisation
    /// with CRM. See TxState.
    /// </summary>
    public enum TransmissionState
    {
        /// <summary>
        /// This is a new SyncState object which has not yet been transmitted.
        /// </summary>
        NewFromOutlook,
        /// <summary>
        /// This is a SyncState representing an outlook item which was present when 
        /// Outlook was started.
        /// </summary>
        PresentAtStartup,
        /// <summary>
        /// This is a SyncState object representing a outlook item which has just been 
        /// created from a CRM item.
        /// </summary>
        NewFromCRM,
        /// <summary>
        /// A change has been registered on this SyncState object but it has
        /// not been transmitted.
        /// </summary>
        Pending,
        /// <summary>
        /// This SyncState has been queued for transmission but has not yet been
        /// transmitted.
        /// </summary>
        Queued,
        /// <summary>
        /// The Outlook item associated with this SyncState has been transmitted,
        /// but no confirmation has yet been received that it has been accepted.
        /// </summary>
        Transmitted,
        /// <summary>
        /// The Outlook item associated with this SyncState has been transmitted
        /// and accepted by CRM.
        /// </summary>
        Synced,
        /// <summary>
        /// A state is put into state PendingDeletion if it is not found in CRM at 
        /// one synchronisation run; if it is not found in the subsequent run and is
        /// still in state PendingDeletion, then it should be deleted.
        /// </summary>
        PendingDeletion,
        /// <summary>
        /// The sync state is in an invalid state and should never be synced.
        /// </summary>
        Invalid
    }
}
