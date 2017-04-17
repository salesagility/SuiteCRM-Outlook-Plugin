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
    /// <typeparam name="OutlookItemType">The type of outlook item this proto-item transduces.</typeparam>
    public abstract class ProtoItem<OutlookItemType> : ProtoItem
    {
    }
}