namespace SuiteCRMAddIn.Exceptions
{
    using BusinessLogic;
    using System;

    /// <summary>
    /// An exception thrown if an unexpected sync state class is found.
    /// </summary>
    [Serializable]
    internal class UnexpectedSyncStateClassException : Exception
    {

        public UnexpectedSyncStateClassException(string expectedClassName, SyncState found) : 
            base($"Unexpected sync state: expected {expectedClassName}, found {found.GetType().Name}")
        {
        }
    }
}