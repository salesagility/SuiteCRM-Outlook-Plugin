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
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;

    public class EntryValue
    {
        /// <summary>
        /// A map of my names/values.
        /// </summary>
        private Dictionary<string, object> map;

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
                this.nameValueList = new NameValueCollection(value);
                this.map = this.nameValueList.AsDictionary();
            }
        }
        public NameValueCollection nameValueList { get; set; }

        public object GetValue(string key)
        {
            object result = null;

            try
            {
                if (this.map != null && this.map.ContainsKey(key))
                {
                    result = this.map[key];
                }
            }
            catch (Exception)
            {
            }

            return result;
        }

        /// <summary>
        /// Get the binding for this name within this entry.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns>The binding.</returns>
        public NameValue GetBinding(string name)
        {
            return this.nameValueList.GetBinding(name);
        }

        public RelationshipListElement relationships { get; set; }

        public string GetValueAsString(string key)
        {
            object value = this.GetValue(key);
            return value == null ? string.Empty : value.ToString();
        }

        /// <summary>
        /// Get the value of the stated key, presumed to be a date/time string, as a date time object
        /// in UTC.
        /// </summary>
        /// <param name="key">The key to seek</param>
        /// <returns>The date/time value in UTC, if it was a date/time value; otherwise, DateTime.MinValue.</returns>
        public DateTime GetValueAsUTC(string key)
        {
            string stringValue = this.GetValueAsString(key);
            DateTime result = DateTime.MinValue;

            if (!string.IsNullOrEmpty(stringValue))
            {
                if (!DateTime.TryParseExact(stringValue, "yyyy-MM-dd HH:mm:ss", null, DateTimeStyles.None, out result))
                {
                    DateTime.TryParse(stringValue, out result);
                }
            }

            /* correct for offset from UTC */
            return result;
        }

        /// <summary>
        /// Get the value of the stated key, presumed to be a date/time string, as a date time object
        /// in local time (the time is delivered by CRM in UTC).
        /// </summary>
        /// <param name="key">The key to seek</param>
        /// <returns>The date/time value in local time, if it was a date/time value; otherwise, DateTime.MinValue.</returns>
        public DateTime GetValueAsDateTime(string key)
        {
            var asUTC = this.GetValueAsUTC(key);

            /* if result is valid, correct for offset from UTC */
            return asUTC == DateTime.MinValue ? asUTC : asUTC.Add(new DateTimeOffset(DateTime.Now).Offset);
        }
    }
}
