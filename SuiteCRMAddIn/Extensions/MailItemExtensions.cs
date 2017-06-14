
namespace SuiteCRMAddIn.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Extension methods for Outlook MailItem objects.
    /// </summary>
    public static class MailItemExtensions
    {
        /// <summary>
        /// Magic property tag to get the email address from an Outlook Recipient object.
        /// </summary>
        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        /// <summary>
        /// From this mail item, extract the SMTP sender address if any, else the
        /// empty string.
        /// </summary>
        /// <remarks>
        /// <para>
        /// If the mail item has not yet been despatched (which in the case of a "Send and 
        /// Archive" action it will not have been), the sender address will not have
        /// been filled in. Currently using the default sender address but I'm not convinced
        /// this is satisfactory.
        /// </para>
        /// <para>
        /// If the sender is using Exchange (which if they're using Outlook they almost
        /// certainly are) the 'sender email address' won't be an email address, it will
        /// be a bizarre LDAP query which CRM will barf on. So resolve it if possible.
        /// </para>
        /// </remarks>
        /// <param name="olItem">The mail item</param>
        /// <returns>An SMTP address or an empty string.</returns>
        public static string GetSenderSMTPAddress(this Outlook.MailItem olItem)
        {
            string result = string.Empty;

            try
            {
                switch (olItem.SenderEmailType)
                {
                    case "SMTP": /* an SMTP address; easy */
                        result = olItem.SenderEmailAddress;
                        break;
                    case "EX": /* an Exchange address */
                        var sender = olItem.Sender;
                        if (sender != null)
                        {
                            var exchangeUser = sender.GetExchangeUser();
                            if (exchangeUser != null)
                            {
                                result = exchangeUser.PrimarySmtpAddress;
                            }
                        }
                        break;
                    case "":
                    case null:
                        /* happens, is coped with in the final clause, don't worry about it */
                        break;
                    default:
                        Globals.ThisAddIn.Log.AddEntry($"Unknown email type {olItem.SenderEmailType}", SuiteCRMClient.Logging.LogEntryType.Warning);
                        break;
                }
            }
            catch (Exception any)
            {
                Globals.ThisAddIn.Log.AddEntry(
                    $"MailItemExtensions.GetSenderSMTPAddress: unexpected error {any.GetType().Name} '{any.Message}'", 
                    SuiteCRMClient.Logging.LogEntryType.Error);
            }

            try
            {
                if (string.IsNullOrEmpty(result))
                {
                    var currentUser = Globals.ThisAddIn.Application.ActiveExplorer().Session.CurrentUser.PropertyAccessor;
                    result = currentUser.GetProperty(PR_SMTP_ADDRESS).ToString();
                }
            }
            catch (Exception any)
            {
                Globals.ThisAddIn.Log.AddEntry(
                    $"MailItemExtensions.GetSenderSMTPAddress: failed to get email address of current user: {any.GetType().Name} '{any.Message}'",
                    SuiteCRMClient.Logging.LogEntryType.Error);
            }

            return result;
        } 
    }
}
