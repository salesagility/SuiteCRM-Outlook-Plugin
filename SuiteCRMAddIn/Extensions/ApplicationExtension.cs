
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
            return app.Session.CurrentUser.GetSmtpAddress();
        }
    }
}
