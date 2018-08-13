/**
 * Outlook integration for SuiteCRM.
 * @package Outlook integration for SuiteCRM
 * @copyright SalesAgility Ltd http://www.salesagility.com
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU LESSER GENERAL PUBLIC LICENCE as published by
 * the Free Software Foundation; either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU LESSER GENERAL PUBLIC LICENCE
 * along with this program; if not, see http://www.gnu.org/licenses
 * or write to the Free Software Foundation,Inc., 51 Franklin Street,
 * Fifth Floor, Boston, MA 02110-1301  USA
 *
 * @author SalesAgility <info@salesagility.com>
 */
namespace SuiteCRMClient.RESTObjects
{
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// A collection of names and values, implemented as a list of name/value objects.
    /// </summary>
    /// <see cref="NameValue"/> 
    public class NameValueCollection : List<NameValue>
    {

        /// <summary> 
        /// It appears that CRM sends us back strings HTML escaped. 
        /// </summary> 
        private JsonSerializerSettings deserialiseSettings = new JsonSerializerSettings()
        {
            StringEscapeHandling = StringEscapeHandling.EscapeHtml
        };

        /// <summary>
        /// Construct a new unpopulated instance of NameValueCollection.
        /// </summary>
        public NameValueCollection()
        {
        }

        /// <summary>
        /// Construct a new instance of NameValueCollection initialised with this `data`.
        /// </summary>
        /// <param name="data">The data with which I should be populated.</param>
        public NameValueCollection(JObject data)
        {
            foreach (object objField in data.ToArray<object>())
            {
                string strFieldString = objField.ToString();
                strFieldString = strFieldString.Remove(0, strFieldString.IndexOf('{'));
                NameValue objActualField = JsonConvert.DeserializeObject<NameValue>(strFieldString, deserialiseSettings);
                this.Add(objActualField);
            }
        }

        /// <summary>
        /// Get the binding for this name within this name-value collection.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns>The binding.</returns>
        public NameValue GetBinding(string name)
        {
            return this.Where(x => x.name == name).FirstOrDefault();
        }

        /// <summary>
        /// Get the value for this name within this name-value collection, as a string.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns>The value as a string (or the empty string if unbound).</returns>
        public string GetValueAsString(string name)
        {
            var binding = this.GetBinding(name);
            return binding == null ?
                string.Empty :
                binding.value.ToString();
        }

        /// <summary>
        /// Return my names/values as a dictionary.
        /// </summary>
        /// <returns>my names/values as a dictionary</returns>
        public Dictionary<string, object> AsDictionary()
        {
            return this.Where(x => !string.IsNullOrEmpty(x.name)).ToDictionary(x => x.name, x => x.value);
        }
    }
}
