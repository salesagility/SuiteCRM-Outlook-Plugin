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
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace SuiteCRMClient.RESTObjects
{
    using Newtonsoft.Json;

    public class Field
    {
        [JsonProperty("default_value")]
        public string default_value { get; set; }
        [JsonProperty("group")]
        public string group { get; set; }
        [JsonProperty("label")]
        public string label { get; set; }
        [JsonProperty("name")]
        public string name { get; set; }

        private object rawOptions;

        [JsonProperty("options")]
        public object optionsField
        {
            get
            {
                return this.rawOptions;
            }
            set
            {
                this.rawOptions = value;

                if (value is JArray)
                {
                    // this means there are no options; do nothing
                }
                else
                {
                    foreach (KeyValuePair<string, JToken> entry in ((JObject)value))
                    {
                        this.Options[entry.Key] = entry.Value.ToObject<NameValue>();
                    }
                }
            }
        }
        public Dictionary<string,RESTObjects.NameValue> Options = new Dictionary<string, NameValue>();

        [JsonProperty("required")]
        public int required { get; set; }
        [JsonProperty("type")]
        public string type { get; set; }

        /// <summary>
        /// ?The name of the link table? Present if type = 'link'
        /// </summary>
        [JsonProperty("relationship")]
        public string relationship { get; set; }

        /// <summary>
        /// The module related to. Present if type = 'link'
        /// </summary>
        [JsonProperty("module")]
        public string module { get; set; }
    }



}
