﻿/**
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
    public class eModuleFields
    {
        [JsonProperty("error")]
        public eErrorValue error { get; set; }

        private JObject module_fieldsField;
        [JsonProperty("module_fields")]

        public JObject module_fields_object
        {
            get
            {
                return this.module_fieldsField;
            }
            set
            {
                this.module_fieldsField = value;
                this.moduleFields = new List<eField>();               
                foreach (object objField in value.ToArray<object>())
                {
                    string fieldSpecification = objField.ToString();
                    fieldSpecification = fieldSpecification.Remove(0, fieldSpecification.IndexOf('{'));
                    eField field = JsonConvert.DeserializeObject<eField>(fieldSpecification);
                    this.moduleFields.Add(field);
                }
            }
        }

        public List<eField> moduleFields { get; set; }

        private JObject link_fieldsField;
        [JsonProperty("link_fields")]

        public JObject link_fields_object
        {
            get
            {
                return this.link_fieldsField;
            }
            set
            {
                this.link_fieldsField = value;
                this.linkFields = new List<eField>();
                foreach (object objField in value.ToArray<object>())
                {
                    string fieldSpecification = objField.ToString();
                    fieldSpecification = fieldSpecification.Remove(0, fieldSpecification.IndexOf('{'));
                    eField field = JsonConvert.DeserializeObject<eField>(fieldSpecification);
                    this.linkFields.Add(field);
                }
            }
        }

        public List<eField> linkFields { get; set; }

        public List<eField> fields
        {
            get
            {
                List<eField> result = new List<eField>();
                result.AddRange(this.moduleFields);
                result.AddRange(this.linkFields);
                return result;
            }
        }

        [JsonProperty("module_name")]
        public string module_name { get; set; }
    }
}
