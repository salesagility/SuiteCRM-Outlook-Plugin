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
namespace SuiteCRMClient.Exceptions
{
    using System;
    using SuiteCRMClient.RESTObjects;

    /// <summary>
    /// An exception which wraps an ErrorValue object.
    /// </summary>
    [Serializable]
    public class CrmServerErrorException : Exception
    {
        /// <summary>
        /// The error as returned by CRM over the JSON link.
        /// </summary>
        public readonly ErrorValue Error;

        /// <summary>
        /// The error number
        /// </summary>
        public int ErrorNumber => Int32.Parse(Error.number);

        /// <summary>
        /// The payload of the request which resulted in the error.
        /// </summary>
        public readonly string payload;

        /// <summary>
        /// Construct a new instance of CrmServerErrorException.
        /// </summary>
        /// <param name="error">The CRM error to wrap.</param>
        public CrmServerErrorException(ErrorValue error) : base($"CRM Server error {error.number} ({error.name}): {error.description}")
        {
            this.Error = error;
        }

        /// <summary>
        /// Construct a new instance of CrmServerErrorException.
        /// </summary>
        /// <param name="error">The CRM error to wrap.</param>
        /// <param name="payload">The payload of the request which resulted in the error.</param>
        public CrmServerErrorException(ErrorValue error, string payload) : base($"CRM Server error {error.number} ({error.name}): {error.description}; request payload: {payload}")
        {
            this.payload = payload;
        }
    }
}