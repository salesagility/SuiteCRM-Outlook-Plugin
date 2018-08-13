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

    public class EntryList
    {
        [JsonProperty("entry_list")]
        public EntryValue[] entry_list { get; set; }

        private RelationshipListElement[] relationshipList;

        /// <summary>
        /// Related records; the result of a `link_names_to_fields_array` clause in the query.
        /// </summary>
        [JsonProperty("relationship_list")]
        public RelationshipListElement[] relationship_list
        {
            get
            {
                return this.relationshipList;
            }
            set
            {
                this.relationshipList = value;
            }
        }


        [JsonProperty("error")]
        public ErrorValue error { get; set; }
        [JsonProperty("field_list")]
        public Field[] field_list { get; set; }
        [JsonProperty("next_offset")]
        public int next_offset { get; set; }
        [JsonProperty("result_count")]
        public int result_count { get; set; }

        public void ResolveLinks()
        {
            if (this.entry_list.Length == this.relationshipList.Length)
            {
                for (int i = 0; i < this.entry_list.Length; i++)
                {
                    this.entry_list[i].relationships = this.relationshipList[i];
                }
            }
        }
    }
}
