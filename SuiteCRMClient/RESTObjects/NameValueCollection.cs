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
    using System.Collections.Generic;

	/// <summary>
	/// A collection of names and values, implemented as a list of name/value objects.
	/// </summary>
	/// <see cref="eNameValue"/> 
    public class NameValueCollection : List<eNameValue>
    {
        /// <summary>
        /// Return my names/values as a dictionary.
        /// </summary>
        /// <returns>my names/values as a dictionary</returns>
        public Dictionary<string, object> AsDictionary()
        {
            Dictionary<string, object> result = new Dictionary<string, object>();

            foreach (eNameValue entry in this)
            {
                result[entry.name] = entry.value;
            }

            return result;
        }
    }
}
