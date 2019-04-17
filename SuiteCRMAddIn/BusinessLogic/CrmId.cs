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

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using SuiteCRMAddIn.Exceptions;
using SuiteCRMAddIn.Properties;

namespace SuiteCRMAddIn.BusinessLogic
{
    /// <summary>
    ///     A validated CRM id.
    /// </summary>
    public class CrmId : IComparable
    {
        public static readonly CrmId Empty = new CrmId();

        private static readonly Regex Validator =
            CrmIdValidationPolicy.GetValidationPattern(Settings.Default.CrmIdValidationPolicy);

        private static readonly Dictionary<string, CrmId> Issued = new Dictionary<string, CrmId>();

        /// <summary>
        ///     The actual id string.
        /// </summary>
        /// <remarks>
        ///     This class would specialise <see cref="string" /> only you can't.
        /// </remarks>
        private readonly string crmId;

        private CrmId()
        {
            crmId = string.Empty;
        }

        /// <summary>
        ///     Create a new instance of a CrmId with this id.
        /// </summary>
        /// <remarks>
        ///     This has to be public so that the JSON deserialiser can use it - but don't use it
        ///     otherwise
        /// </remarks>
        /// <param name="id"></param>
        public CrmId(string id)
        {
            if (IsValid(id))
            {
                crmId = id;
                Issued[id] = this;
            }
            else
            {
                throw new TypeInitializationException(GetType().FullName,
                    new InvalidCrmIdException($"'{id}' does not appear to be a valid CRM id."));
            }
        }

        public int CompareTo(object obj)
        {
            return crmId.CompareTo(obj.ToString());
        }

        public override string ToString()
        {
            return crmId;
        }

        /// <summary>
        ///     Validates a CRM id.
        /// </summary>
        /// <param name="id">The string which may or may not be a valid CRM id.</param>
        /// <returns>
        ///     True if `id` matches <see cref="Validator" /> pattern and
        ///     is of suitable length.
        /// </returns>
        public static bool IsValid(string id)
        {
            return !string.IsNullOrEmpty(id) && Validator.IsMatch(id);
        }

        /// <summary>
        ///     Validates a CRM id.
        /// </summary>
        /// <param name="id">The object which may or may not be a valid CRM id.</param>
        /// <returns>
        ///     True if `id` is not null, matches <see cref="Validator" /> pattern and
        ///     is of suitable length.
        /// </returns>
        public static bool IsValid(CrmId id)
        {
            return id != null && id.IsValid();
        }

        /// <summary>
        ///     Validates a CRM id.
        /// </summary>
        /// <returns>
        ///     True if I match <see cref="Validator" /> pattern and
        ///     am of suitable length.
        /// </returns>
        public bool IsValid()
        {
            return IsValid(crmId);
        }

        /// <summary>
        ///     True if <see cref="CrmId.IsValid(CrmId)" /> is false of this id.
        /// </summary>
        /// <param name="id">The object which may or may not be a valid CRM id.</param>
        /// <returns>True if <see cref="CrmId.IsValid(CrmId)" /> is false of this id.</returns>
        public static bool IsInvalid(CrmId id)
        {
            return !IsValid(id);
        }

        public override bool Equals(object obj)
        {
            return base.Equals(obj) || (obj as CrmId)?.ToString() == crmId;
        }

        public override int GetHashCode()
        {
            return crmId.GetHashCode();
        }

        /// <summary>
        ///     Get the single CrmId instance for this value.
        /// </summary>
        /// <param name="value">The value to seek.</param>
        /// <returns>A CrmId instance</returns>
        /// <exception cref="TypeInitializationException"> if `value` does not appear to be a valid CRM id.</exception>
        public static CrmId Get(string value)
        {
            return IsValid(value) ? Issued.ContainsKey(value) ? Issued[value] : new CrmId(value) : Empty;
        }

        /// <summary>
        ///     Get the single CrmId instance for this value.
        /// </summary>
        /// <param name="value">The value to seek.</param>
        /// <returns>A CrmId instance</returns>
        /// <exception cref="TypeInitializationException"> if `value` does not appear to be a valid CRM id.</exception>
        public static CrmId Get(object value)
        {
            return value == null ? Empty : Get(value.ToString());
        }
    }
}