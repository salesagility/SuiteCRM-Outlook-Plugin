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
namespace SuiteCRMClient.Email
{
    using System.Collections.Generic;
    using System.Linq;

    public class ArchiveResult
    {
        public static ArchiveResult Success(string emailId, IEnumerable<System.Exception> warnings)
        {
            return new ArchiveResult
            {
                EmailId = emailId,
                Problems = warnings,
            };
        }

        public static ArchiveResult Failure(params System.Exception[] exceptions)
        {
            return new ArchiveResult
            {
                Problems = exceptions,
            };
        }

        public string EmailId { get; set; }

        public IEnumerable<System.Exception> Problems { get; set; }

        public bool IsSuccess => !string.IsNullOrEmpty(EmailId) && (Problems == null || Problems.Count() == 0);

        public bool IsFailure => !IsSuccess;
    }
}
