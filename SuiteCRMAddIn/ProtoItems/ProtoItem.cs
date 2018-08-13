namespace SuiteCRMAddIn.ProtoItems
{
    using SuiteCRMClient.RESTObjects;

    /// <summary>
    /// A C# proxy for a CRM object.
    /// </summary>
    /// <remarks>
    /// ProtoItems have two jobs: one is acting as a cache for values read from a 
    /// CRM item, and the other is to act as a transducer between CRM items and 
    /// Outlook items.
    /// </remarks>
    public abstract class ProtoItem
    {
        /// <summary>
        /// Construct a name-value collection from my fields, suitable to be despatched to CRM
        /// to create or update the representation in CRM of the item I represent.
        /// </summary>
        /// <returns>The name-value collection constructed.</returns>
        public abstract NameValueCollection AsNameValues();

        /// <summary>
        /// Get a description of the item.
        /// </summary>
        public abstract string Description { get; }
    }
}
