namespace SuiteCRMAddIn.BusinessLogic
{
    using Newtonsoft.Json;

    public class LicenceValidation
    {
        /// <summary>
        /// true if this LicenceValidation represents a validated licence, else false.
        /// </summary>
        [JsonProperty("validated")]
        public bool validated { get; set; }
    }
}