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
    using Microsoft.Office.Interop.Outlook;
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
        public static string GetSMTPEmailAddress(MailItem mailItem)
        {
            string str2;
            string str = string.Empty;
            if (((str2 = Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder.Name) != null) && (str2 == "Sent Items"))
            {
                foreach (Recipient recipient in mailItem.Recipients)
                {
                    if (recipient.AddressEntry.Type == "EX")
                    {
                        str = str + GetEmailAddressForExchangeServer(recipient.AddressEntry.Name) + ",";
                    }
                    else
                    {
                        str = str + recipient.Address + ",";
                    }
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
                MailItem item = (MailItem)Globals.ThisAddIn.Application.ActiveExplorer().Application.CreateItem(OlItemType.olMailItem);
                Recipient recipient = item.Recipients.Add(emailName);
                recipient.Resolve();
                ExchangeUser exchangeUser = recipient.AddressEntry.GetExchangeUser();
                if (exchangeUser.PrimarySmtpAddress != string.Empty)
                {
                    return exchangeUser.PrimarySmtpAddress;
                }
                return exchangeUser.PrimarySmtpAddress;
            }
            catch (System.Exception)
            {
                // Swallow exception(!)
                return string.Empty;
            }
        }
        public static string GetSenderAddress(MailItem mail, string type)
        {
            if (type == "SendArchive")
            {
                if (Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry.Type == "EX")
                {
                    return GetEmailAddressForExchangeServer(Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry.Name);
                }
                return Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry.Address;
            }
            if (mail.SenderEmailType == "EX")
            {
                return GetEmailAddressForExchangeServer(mail.SenderName);
            }
            return mail.SenderEmailAddress;
        }
    }
}
