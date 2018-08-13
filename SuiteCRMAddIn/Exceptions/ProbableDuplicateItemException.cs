namespace SuiteCRMAddIn.Exceptions
{
    using System;
    using System.Runtime.Serialization;

    [Serializable]
    internal class ProbableDuplicateItemException<ItemType> : Exception
    {
        private ItemType olItem;
        private string v;

        public ProbableDuplicateItemException()
        {
        }

        public ProbableDuplicateItemException(string message) : base(message)
        {
        }

        public ProbableDuplicateItemException(string message, Exception innerException) : base(message, innerException)
        {
        }

        public ProbableDuplicateItemException(ItemType olItem, string v)
        {
            this.olItem = olItem;
            this.v = v;
        }

        protected ProbableDuplicateItemException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}