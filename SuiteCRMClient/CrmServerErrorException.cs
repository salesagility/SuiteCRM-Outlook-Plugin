using System;
using System.Runtime.Serialization;
using SuiteCRMClient.RESTObjects;

namespace SuiteCRMClient
{
    /// <summary>
    /// An exception which wraps an eErrorValue object.
    /// </summary>
    [Serializable]
    internal class CrmServerErrorException : Exception
    {
        public readonly ErrorValue Error;

        public CrmServerErrorException(ErrorValue error) : base($"CRM Server error {error.number} ({error.name}): {error.description}")
        {
            this.Error = error;
        }
    }
}