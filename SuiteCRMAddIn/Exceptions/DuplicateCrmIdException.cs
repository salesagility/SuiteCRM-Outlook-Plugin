using SuiteCRMAddIn.BusinessLogic;
using System;
using System.Runtime.Serialization;

namespace SuiteCRMAddIn.Exceptions
{
    [Serializable]
    internal class DuplicateCrmIdException : Exception
    {

        public DuplicateCrmIdException()
        {
        }

        public DuplicateCrmIdException(string message) : base(message)
        {
        }

        public DuplicateCrmIdException(SyncState syncState, string id) : base($"Shouldn't happen: more than one Outlook object with CRM id '{id}' ({syncState.Description})")
        {
        }

        public DuplicateCrmIdException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected DuplicateCrmIdException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}