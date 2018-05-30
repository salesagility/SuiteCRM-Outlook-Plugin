using System;
using System.Runtime.Serialization;

namespace SuiteCRMClient
{
    /// <summary>
    /// An exception thrown if/when an attempt is made to send nothing 
    /// (null or an empty name/value array) as a record to CRM.
    /// </summary>
    [Serializable]
    internal class MissingValuesException : Exception
    {
        public MissingValuesException()
        {
        }

        public MissingValuesException(string message) : base(message)
        {
        }

        public MissingValuesException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected MissingValuesException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}