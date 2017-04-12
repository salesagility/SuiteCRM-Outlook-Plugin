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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SuiteCRMClient.RESTObjects
{
    public class eEntryValue
    {
        /// <summary>
        /// A map of my names/values.
        /// </summary>
        private Dictionary<string, object> map = new Dictionary<string, object>();

        /// <summary> 
        /// It appears that CRM sends us back strings HTML escaped. 
        /// </summary> 
        private JsonSerializerSettings deserialiseSettings = new JsonSerializerSettings()
        {
            StringEscapeHandling = StringEscapeHandling.EscapeHtml
        };

        [JsonProperty("id")]
        public string id { get; set; }
        [JsonProperty("module_name")]
        public string module_name { get; set; }
        private JObject name_value_objectField;
        [JsonProperty("name_value_list")]
        public JObject name_value_object
        {
            get
            {
                return this.name_value_objectField;
            }
            set
            {
                this.name_value_objectField = value;
                this.name_value_list1 = new NameValueCollection();
                foreach (object objField in value.ToArray<object>())
                {
                    string strFieldString = objField.ToString();
                    strFieldString = strFieldString.Remove(0, strFieldString.IndexOf('{'));
                    eNameValue objActualField = JsonConvert.DeserializeObject<eNameValue>(strFieldString, deserialiseSettings);
                    this.name_value_list1.Add(objActualField);
                    this.map[objActualField.name] = objActualField.value;
                }
            }
        }
        public NameValueCollection name_value_list1 { get; set; }

        public object GetValue(string key)
        {
            object result = null;

            try
            {
                result = this.map[key];
            }
            catch (Exception)
            {
            }

            return result;
        }

        public string GetValueAsString(string key)
        {
            object value = this.GetValue(key);
            return value == null ? string.Empty : value.ToString();
        }
    }
}
