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

    /// <summary>
    /// An object returned by the get_server_info call.
    /// </summary>
    public class ServerInfo
    {
        /// <summary>
        /// Field returned in the get_server_info packet, but I don yet know what it means.
        /// </summary>
        [JsonProperty("flavor")]
        public string Flavor { get; set; }

        /// <summary>
        /// Version of SugarCRM on which this version of SuiteCRM is based.
        /// </summary>
        [JsonProperty("version")]
        public string SugarVersion { get; set; }

        /// <summary>
        /// Version of SuiteCRM. Note: this field is not present before SuiteCRM version 7.8.5.
        /// </summary>
        [JsonProperty("suitecrm_version")]
        public string SuiteCRMVersion { get; set; }
    }
}
