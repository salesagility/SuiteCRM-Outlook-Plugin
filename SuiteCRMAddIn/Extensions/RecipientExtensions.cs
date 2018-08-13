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
    using System.Collections.Generic;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Extension methods for Outlook Resipient objects.
    /// </summary>
    public static class RecipientExtensions
    {
        /// <summary>
        /// A cache of SMTP addresses, so we're not continually fetching them from a remote 
        /// Exchange server.
        /// </summary>
        private static Dictionary<Outlook.Recipient, string> smtpAddressCache = new Dictionary<Outlook.Recipient, string>();

        /// <summary>
        /// From this email recipient, extract the SMTP address (if that's possible).
        /// </summary>
        /// <param name="recipient">A recipient object</param>
        /// <returns>The SMTP address for that object, if it can be recovered, else an empty string.</returns>
        public static string GetSmtpAddress(this Outlook.Recipient recipient)
        {
            string result = string.Empty;

            if (smtpAddressCache.ContainsKey(recipient))
            {
                result = smtpAddressCache[recipient];
            }

            if (string.IsNullOrEmpty(result))
            {
                switch (recipient.AddressEntry.Type)
                {
                    case "SMTP":
                        result = recipient.Address;
                        break;
                    case "EX": /* an Exchange address */
                        var exchangeUser = recipient.AddressEntry.GetExchangeUser();
                        if (exchangeUser != null)
                        {
                            result = exchangeUser.PrimarySmtpAddress;
                        }
                        break;
                    default:
                        Globals.ThisAddIn.Log.AddEntry(
                            $"RecipientExtensions.GetSmtpAddres: unknown email type {recipient.AddressEntry.Type}",
                            SuiteCRMClient.Logging.LogEntryType.Warning);
                        break;
                }

                smtpAddressCache[recipient] = result;
            }

            return result;
        }


        /// <summary>
        /// Get the token used by CRM to reflect the acceptance status of this recipient.
        /// </summary>
        /// <param name="recipient">The recipient.</param>
        /// <returns>The acceptance status.</returns>
        public static string CrmAcceptanceStatus(this Outlook.Recipient recipient)
        {
            string acceptance = string.Empty;

            switch (recipient.MeetingResponseStatus)
            {
                case Outlook.OlResponseStatus.olResponseAccepted:
                case Microsoft.Office.Interop.Outlook.OlResponseStatus.olResponseOrganized:
                    acceptance = "accept";
                    break;
                case Outlook.OlResponseStatus.olResponseTentative:
                    acceptance = "tentative";
                    break;
                case Microsoft.Office.Interop.Outlook.OlResponseStatus.olResponseNone:
                case Microsoft.Office.Interop.Outlook.OlResponseStatus.olResponseNotResponded:
                    // nothing to do
                    break;
                case Outlook.OlResponseStatus.olResponseDeclined:
                default:
                    acceptance = "decline";
                    break;
            }

            return acceptance;
        }

    }
}
