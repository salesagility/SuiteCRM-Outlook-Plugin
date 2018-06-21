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

    public class SetRelationshipParams
    {
        private string m1;
        private string m2;

        [JsonProperty("module1_id")]
        public string module1_id { get; set; }
        [JsonProperty("module1")]
        public string module1
        {
            get
            {
                return m1;
            }
            set
            {
                m1 = char.ToUpper(value[0]) + value.Substring(1);
            }
        }
        [JsonProperty("module2_id")]
        public string module2_id { get; set; }
        [JsonProperty("module2")]
        public string module2
        {
            get
            {
                return m2;
            }
            set
            {
                m2 = char.ToUpper(value[0]) + value.Substring(1);
            }
        }

        /// <summary>
        /// Only required if you want to delete a relationsip, in which case set it to 1.
        /// </summary>
        [JsonProperty("delete")]
        public int delete { get; set; } = 0;
    }
}
