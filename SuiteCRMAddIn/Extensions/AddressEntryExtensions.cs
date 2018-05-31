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
namespace SuiteCRMAddIn.Extensions
{
    using BusinessLogic;
    using SuiteCRMClient.Logging;
    using System;
    using System.Collections.Generic;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public static class AddressEntryExtensions
    {
        /// <summary>
        /// See https://msdn.microsoft.com/en-us/library/office/ff184624.aspx?f=255&MSPPError=-2147217396
        /// </summary>
        private static string PR_SMTP_ADDRESS = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        /// <summary>
        /// A cache of SMTP addresses, so we're not continually fetching them from a remote 
        /// Exchange server.
        /// </summary>
        private static Dictionary<Outlook.AddressEntry, string> smtpAddressCache = new Dictionary<Outlook.AddressEntry, string>();

        /// <summary>
        /// Shorthand to refer to the global log.
        /// </summary>
        public static ILogger Log
        {
            get
            {
                return Globals.ThisAddIn.Log;
            }
        }

        /// <summary>
        /// Get the SMTP address implied by this AddressEntry object
        /// </summary>
        /// <remarks>
        /// This is different from RecipientExtension.GetSmtpAddress because
        /// we don't have access to anything equivalent to a Recipient object.
        /// </remarks>
        /// <see cref="RecipientExtensions.GetSmtpAddress(Outlook.Recipient)"/> 
        /// <param name="entry">the AddressEntry</param>
        /// <returns>The SMTP address, if it can be recovered, else the empty string.</returns>
        public static string GetSmtpAddress(this Outlook.AddressEntry entry)
        {
            string result;

            try
            {
                result = smtpAddressCache[entry];
            }
            catch (KeyNotFoundException)
            {
                result = string.Empty;
            }

            if (string.IsNullOrEmpty(result))
            {
                try
                {
                    switch (entry.AddressEntryUserType)
                    {
                        case Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry:
                        case Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry:
                            Outlook.ExchangeUser exchUser = entry.GetExchangeUser();
                            result = exchUser == null ?
                                string.Empty :
                                exchUser.PrimarySmtpAddress;
                            break;
                        default:
                            result = entry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                            break;
                    }
                }
                catch (Exception any)
                {
                    ErrorHandler.Handle("Failed while trying to obtain an SMTP address", any);
                }
            }
            smtpAddressCache[entry] = result;

            return result;
        }
    }
}