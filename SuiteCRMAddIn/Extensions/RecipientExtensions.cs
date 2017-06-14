namespace SuiteCRMAddIn.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Extension methods for Outlook Resipient objects.
    /// </summary>
    public static class RecipientExtensions
    {
        /// <summary>
        /// From this email recipient, extract the SMTP address (if that's possible).
        /// </summary>
        /// <param name="recipient">A recipient object</param>
        /// <returns>The SMTP address for that object, if it can be recovered, else an empty string.</returns>
        public static string GetSmtpAddress(this Outlook.Recipient recipient)
        {
            string result = string.Empty;

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

            return result;
        }

    }
}
