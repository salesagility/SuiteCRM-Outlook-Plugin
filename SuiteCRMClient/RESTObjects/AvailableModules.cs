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

    public class AvailableModules
    {
        [JsonProperty("error")]
        public ErrorValue error { get; set; }

        //[JsonProperty("modules")]
        public List<AvailableModule> items { get; set; }

        private List<JObject> module_fieldsField;
        [JsonProperty("modules")]
        public List<JObject> module_fields_object
        {
            get
            {
                return this.module_fieldsField;
            }
            set
            {
                this.module_fieldsField = value;
                this.items = new List<AvailableModule>();
                foreach (object objField in value.ToArray<object>())
                {
                    string strFieldString = objField.ToString();
                    AvailableModule objActualField = JsonConvert.DeserializeObject<AvailableModule>(strFieldString);
                    this.items.Add(objActualField);
                }
            }
        }

        private List<Field> modules { get; set; }

    }

    public class AvailableModule
    {
        [JsonProperty("module_key")]
        public string module_key { get; set; }
        [JsonProperty("module_label")]
        public string module_label { get; set; }

        public List<module_access> module_acls1 { get; set; }

        private List<JObject> _module_acls;
        [JsonProperty("acls")]
        public List<JObject> module_acls
        {
            get
            {
                return this._module_acls;
            }
            set
            {
                this._module_acls = value;
                this.module_acls1 = new List<module_access>();
                foreach (object objField in value.ToArray<object>())
                {
                    string strFieldString = objField.ToString();
                    module_access objActualField = JsonConvert.DeserializeObject<module_access>(strFieldString);
                    this.module_acls1.Add(objActualField);
                }
            }
        }
    }

    public class module_access
    {
        public string action { get; set; }
        public bool access { get; set; }
    }
}
