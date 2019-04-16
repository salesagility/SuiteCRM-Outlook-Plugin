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

#region

using System.Collections.Generic;
using System.Linq;
using System.Text;
using SuiteCRMAddIn.BusinessLogic;
using SuiteCRMAddIn.Exceptions;
using SuiteCRMClient;
using SuiteCRMClient.RESTObjects;

#endregion

namespace SuiteCRMAddIn.Helpers
{
    /// <summary>
    ///     We do search (differently) in too many places. This is an attempt to rationalise it.
    /// </summary>
    public class SearchHelper
    {
        public static IEnumerable<EntryValue> SearchContacts(string token)
        {
            return Search(ContactSynchroniser.CrmModule, token,
                new[] {"first_name", "last_name", "email1" , "sync_contact", "outlook_id" });
        }

        public static IEnumerable<EntryValue> Search(string module, string token, IEnumerable<string> fields,
            string logicalOperator = "OR")
        {
            var bob = new StringBuilder("(");
            var fieldsArray = fields.ToArray();

            foreach (var field in fieldsArray)
            {
                switch (field)
                {
                    case "first_name":
                    case "last_name":
                    case "name":
                        if (field != fieldsArray.First())
                            bob.Append($"{logicalOperator} ");
                        bob.Append($"{module.ToLower()}.{field} ").Append(token.Length < 4
                            ? $"= '{token}' "
                            : $"LIKE '%{token}%' ");
                        break;
                    case "email1":
                        if (field != fieldsArray.First())
                            bob.Append($"{logicalOperator} ");
                        bob.Append(
                            $"({module.ToLower()}.id in (select eabr.bean_id from email_addr_bean_rel eabr INNER JOIN email_addresses ea on eabr.email_address_id = ea.id  where eabr.bean_module = '{module}' and ea.email_address ");
                        bob.Append(token.Length < 4 ? $"= '{token}'))" : $"LIKE '%{token}%'))");
                        break;
                }
            }
            bob.Append(")");

            var result = RestAPIWrapper.GetEntryList(module, bob.ToString(), 1000, "date_entered DESC", 0, false, fieldsArray)
                .entry_list;

            return result;
        }
    }
}
