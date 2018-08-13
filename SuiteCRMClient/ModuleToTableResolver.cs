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
namespace SuiteCRMClient
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// In general the module key for a module (that is, the standard name for the module, generally the 
    /// same as the American English version of the name) is the singular (i.e. without the terminal 's') 
    /// of the name of the table in which the module data is stored. There are a few exceptions to this.
    /// </summary>
    public class ModuleToTableResolver
    {
        /// <summary>
        /// the overrides to the general rule. This probably ought to be read from a setting at startup.
        /// </summary>
        private static Dictionary<string, string> overrides = new Dictionary<string, string>
        {
            { "Projects", "Project" },
            { "Project", "Project" }
        };

        /// <summary>
        /// Return the table name which corresponds to this module name.
        /// </summary>
        /// <param name="moduleName">The module name.</param>
        /// <returns>The corresponding table name.</returns>
        public static string GetTableName(string moduleName)
        {
            string result;

            try
            {
                result = overrides[moduleName];
            }
            catch (Exception)
            {
                result = moduleName.EndsWith("s") ? moduleName : $"{moduleName}s";
            }

            return result;
        }
    }
}
