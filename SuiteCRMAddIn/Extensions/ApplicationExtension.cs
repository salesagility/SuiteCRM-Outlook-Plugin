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
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using BusinessLogic;
    using SuiteCRMClient;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System.Runtime.InteropServices;
    using SuiteCRMClient.Logging;

    public static class ApplicationExtension
    {
        /// <summary>
        /// Get a best approximation of the username of the current user.
        /// </summary>
        /// <returns>a best approximation of the username of the current user</returns>
        public static string GetCurrentUsername(this Outlook.Application app)
        {
            string result;
            Outlook.Recipient currentUser = app.Session.CurrentUser;
            Outlook.AddressEntry addressEntry = currentUser.AddressEntry;

            if (addressEntry.Type == "EX")
            {
                result = addressEntry.GetExchangeUser().Name;
            }
            else
            {

                result = currentUser.Name;
                // result = addressEntry.Address.Substring(0, addressEntry.Address.IndexOf('@'));
            }

            return result;
        }

        /// <summary>
        /// Extract the SMTP address of the current user (if that's possible).
        /// </summary>
        /// <returns>The SMTP address for the current user, if it can be recovered, else an empty string.</returns>
        public static string GetCurrentUserSMTPAddress(this Outlook.Application app)
        {
            string result;

            try
            {
                result = app.Session.CurrentUser.GetSmtpAddress();
            }
            catch (COMException any)
            {
                // this happens, I think when there's been a failure to connect to Exchange server.
                Globals.ThisAddIn.Log.Error("Failure when trying to access user SMTP address", any);
                result = string.Empty;
            }

            return result;
        }
    }
}
