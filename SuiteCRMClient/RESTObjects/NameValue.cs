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
    using System;
    using System.Globalization;
    using System.Web;

    public class NameValue
    {
        private object v;

        [JsonProperty("name")]
        public string name { get; set; }

        [JsonProperty("value")]
        public object value
        {
            get
            {
                object result = v;
                /// There's a problem with HTML decoding things which come back over the JSON link.
                /// This is a hack; it's almost certainly not the best solution. It also doesn't work.
                if (v is string || v is String)
                {
                    string sv = v.ToString();
                    string decode = HttpUtility.HtmlDecode(sv);
                    if (!decode.Equals(sv))
                    {
                         result = decode;
                    }
                }

                return result;
            }
            set
            {
                // v = value;
                // This might be better: 
                v = value == null ? string.Empty : value;
            }
        }

        /// <summary>
        /// Return the value of this object as a DateTime, if it can be represented as such, otherwise DateTime.MinValue
        /// </summary>
        /// <returns>the value of this object as a DateTime, if it can be represented as such, otherwise DateTime.MinValue</returns>
        public DateTime AsDateTime()
        {
            var stringValue = this.value.ToString();
            DateTime result = DateTime.MinValue;

            if (!DateTime.TryParseExact(stringValue, "yyyy-MM-dd HH:mm:ss", null, DateTimeStyles.None, out result))
            {
                DateTime.TryParse(stringValue, out result);
            }

            return result;
        }
    }

}
