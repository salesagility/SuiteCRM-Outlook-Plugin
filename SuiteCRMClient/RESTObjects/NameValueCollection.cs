namespace SuiteCRMClient.RESTObjects
{
    using System.Collections.Generic;

    public class NameValueCollection : List<eNameValue>
    {
        /// <summary>
        /// Return my names/values as a dictionary.
        /// </summary>
        /// <returns>my names/values as a dictionary</returns>
        public Dictionary<string, object> AsDictionary()
        {
            Dictionary<string, object> result = new Dictionary<string, object>();

            foreach (eNameValue entry in this)
            {
                result[entry.name] = entry.value;
            }

            return result;
        }
    }
}
