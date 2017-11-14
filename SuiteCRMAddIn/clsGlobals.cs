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
namespace SuiteCRMAddIn
{
    using Extensions;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System.Text.RegularExpressions;


    public static class clsGlobals
    {
      
        public static string MySqlEscape(string usString)
        {
            if (usString == null)
            {
                return null;
            }
            return Regex.Replace(usString, "[\\r\\n\\x00\\x1a\\\\'\"]", @"\$0");
        }
        public static string GetSMTPEmailAddress(Outlook.MailItem mailItem)
        {
            string str2;
            string str = string.Empty;
            if (((str2 = Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder.Name) != null) && (str2 == "Sent Items"))
            {
                foreach (Outlook.Recipient recipient in mailItem.Recipients)
                {
                    str += $"{recipient.GetSmtpAddress()},";
                }
            }
            else if (mailItem.SenderEmailType == "EX")
            {
                str = GetEmailAddressForExchangeServer(mailItem.SenderName) + ";";
            }
            else
            {
                str = str + mailItem.SenderEmailAddress + ";";
            }
            return str.Remove(str.Length - 1, 1);
        }

        public static string GetEmailAddressForExchangeServer(string emailName)
        {
            try
            {
                Outlook.MailItem item = (Outlook.MailItem)Globals.ThisAddIn.Application.ActiveExplorer().Application.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Recipient recipient = item.Recipients.Add(emailName);
                recipient.Resolve();
                Outlook.ExchangeUser exchangeUser = recipient.AddressEntry.GetExchangeUser();
                return exchangeUser == null ? string.Empty : exchangeUser.PrimarySmtpAddress;
            }
            catch (System.Exception)
            {
                // Swallow exception(!)
                return string.Empty;
            }
        }
        public static string GetSenderAddress(Outlook.MailItem mail, string type)
        {
            if (type == "SendArchive")
            {
                var addressEntry = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
                if (addressEntry.Type == "EX")
                {
                    return GetEmailAddressForExchangeServer(addressEntry.Name);
                }
                return addressEntry.Address;
            }
            if (mail.SenderEmailType == "EX")
            {
                return GetEmailAddressForExchangeServer(mail.SenderName);
            }
            return mail.SenderEmailAddress;
        }


        /// <summary>
        /// Get a best approximation of the username of the current user.
        /// </summary>
        /// <returns>a best approximation of the username of the current user</returns>
        public static string GetCurrentUsername()
        {
            string result;
            Outlook.AddressEntry addressEntry = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
            if (addressEntry.Type == "EX")
            {
                Outlook.ExchangeUser currentUser =
                    Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry.GetExchangeUser();

                result = currentUser.Name;
            }
            else
            {
                result = addressEntry.Address.Substring(addressEntry.Address.IndexOf('@'));
            }

            return result;
        }
    }
}
