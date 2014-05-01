/**
 * Outlook integration for SuiteCRM.
 * @package Outlook integration for SuiteCRM
 * @copyright SalesAgility Ltd http://www.salesagility.com
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU AFFERO GENERAL PUBLIC LICENSE as published by
 * the Free Software Foundation; either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU AFFERO GENERAL PUBLIC LICENSE
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

namespace SuiteCRMClient.RESTObjects
{
    public class eField
    {
        [JsonProperty("default_value")]
        public string default_value { get; set; }
        [JsonProperty("group")]
        public string group { get; set; }
        [JsonProperty("label")]
        public string label { get; set; }
        [JsonProperty("name")]
        public string name { get; set; }
        //[JsonProperty("options")]
        //public List<RESTObjects.name_value> options { get; set; }
        [JsonProperty("required")]
        public int required { get; set; }
        [JsonProperty("type")]
        public string type { get; set; }
    }



}
