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
                this.name_value_list1 = new List<eNameValue>();
                foreach (object objField in value.ToArray<object>())
                {
                    string strFieldString = objField.ToString();
                    strFieldString = strFieldString.Remove(0, strFieldString.IndexOf('{'));
                    eNameValue objActualField = JsonConvert.DeserializeObject<eNameValue>(strFieldString, deserialiseSettings);
                    this.name_value_list1.Add(objActualField);
                }
            }
        }
        public List<eNameValue> name_value_list1 { get; set; }
    }
}
