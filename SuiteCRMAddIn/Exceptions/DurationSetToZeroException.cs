using System;
using System.Runtime.Serialization;

namespace SuiteCRMAddIn.Exceptions
{
    /// <summary>
    /// #6034: occasionally we get spurious ItemChange events where the 
    /// value of Duration appear as zero, although nothing has occured to
    /// make this change. This is a hack around the problem while we try
    /// to understand it better.
    /// </summary>
    [Serializable]
    internal class DurationSetToZeroException : Exception
    {
        private int duration;

        public int Duration => this.duration;

        public DurationSetToZeroException()
        {
        }

        public DurationSetToZeroException(string message) : base(message)
        {
        }

        public DurationSetToZeroException(int duration)
        {
            this.duration = duration;
        }

        public DurationSetToZeroException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected DurationSetToZeroException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}