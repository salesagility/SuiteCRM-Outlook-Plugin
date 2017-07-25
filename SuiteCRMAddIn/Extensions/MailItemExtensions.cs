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
    using Exceptions;
    using SuiteCRMClient;
    using SuiteCRMClient.Email;
    using SuiteCRMClient.Logging;
    using System;
    using System.Runtime.InteropServices;
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
        /// The name of the magic category we set when a mail is successfully archived.
        /// </summary>
        public const string SuiteCRMCategoryName = "SuiteCRM";

        /// <summary>
        /// The name of the CRM ID synchronisation property.
        /// </summary>
        /// <see cref="SuiteCRMAddIn.BusinessLogic.Synchroniser{OutlookItemType}.CrmIdPropertyName"/> 
        public const string CrmIdPropertyName = "SEntryID";

        /// <summary>
        /// The name of the Outlook user property on which we will store the CRM Category associated
        /// with an email, of any.
        /// </summary>
        public const string CRMCategoryPropertyName = "SuiteCRMCategory";

        /// <summary>
        /// Shorthand to refer to the global user session.
        /// </summary>
        public static UserSession SuiteCRMUserSession
        {
            get
            {
                return Globals.ThisAddIn.SuiteCRMUserSession;
            }
        }

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
                        Log.Warn($"Unknown email type {olItem.SenderEmailType}");
                        break;
                }
            }
            catch (Exception any)
            {
                Log.Error(
                    $"MailItemExtensions.GetSenderSMTPAddress: unexpected error {any.GetType().Name} '{any.Message}'", any);
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
                Log.Error(
                    $"MailItemExtensions.GetSenderSMTPAddress: failed to get email address of current user: {any.GetType().Name} '{any.Message}'",
                    any);
            }

            return result;
        }


        /// <summary>
        /// Constuct an achiveable email object represtenting me.
        /// </summary>
        /// <param name="olItem">Me.</param>
        /// <param name="reason">The reason I should be archived.</param>
        /// <returns>An achiveable email object represtenting me.</returns>
        public static ArchiveableEmail AsArchiveable(this Outlook.MailItem olItem, EmailArchiveReason reason)
        {
            ArchiveableEmail mailArchive = new ArchiveableEmail(MailItemExtensions.SuiteCRMUserSession, MailItemExtensions.Log);
            mailArchive.From = olItem.GetSenderSMTPAddress();
            mailArchive.To = string.Empty;

            Log.Info($"EmailArchiving.SerialiseEmailObject: serialising mail {olItem.Subject} dated {olItem.SentOn}.");

            foreach (Outlook.Recipient recipient in olItem.Recipients)
            {
                string address = recipient.GetSmtpAddress();

                if (mailArchive.To == string.Empty)
                {
                    mailArchive.To = address;
                }
                else
                {
                    mailArchive.To += ";" + address;
                }
            }

            mailArchive.OutlookId = olItem.EnsureEntryID();
            mailArchive.Subject = olItem.Subject;
            mailArchive.Sent = olItem.ArchiveTime(reason);
            mailArchive.Body = olItem.Body;
            mailArchive.HTMLBody = olItem.HTMLBody;
            mailArchive.Reason = reason;
            mailArchive.Category = olItem.UserProperties[CRMCategoryPropertyName] != null ?
                olItem.UserProperties[CRMCategoryPropertyName].Value :
                string.Empty;

            if (Properties.Settings.Default.ArchiveAttachments)
            {
                foreach (Outlook.Attachment attachment in olItem.Attachments)
                {
                    mailArchive.Attachments.Add(new ArchiveableAttachment
                    {
                        DisplayName = attachment.DisplayName,
                        FileContentInBase64String = olItem.GetAttachmentAsBytes(attachment)
                    });
                }
            }

            return mailArchive;
        }


        /// <summary>
        /// An Outlook item doesn't get its entry ID until the first time it's saved; if I don't have one, save me.
        /// </summary>
        /// <param name="olItem">Me</param>
        /// <returns>My entry id.</returns>
        public static string EnsureEntryID(this Outlook.MailItem olItem)
        {
            if (olItem.EntryID == null)
            {
                olItem.Save(); 
            }

            return olItem.EntryID;
        }


        /// <summary>
        /// Get this attachment of mine as an array of bytes.
        /// </summary>
        /// <param name="olItem">Me</param>
        /// <param name="attachment">The attachment to serialise.</param>
        /// <returns>An array of bytes representing the attachment.</returns>
        public static byte[] GetAttachmentAsBytes(this Outlook.MailItem olItem, Outlook.Attachment attachment)
        {
            byte[] result = null;

            Log.Info($"EmailArchiving.GetAttachmentBytes: serialising attachment '{attachment.FileName}' of email '{olItem.Subject}'.");

            if (attachment != null)
            {
                var tempPath = System.IO.Path.GetTempPath();
                string uid = Guid.NewGuid().ToString();
                var temporaryAttachmentPath = $"{tempPath}\\Attachments_{uid}";

                if (!System.IO.Directory.Exists(temporaryAttachmentPath))
                {
                    System.IO.Directory.CreateDirectory(temporaryAttachmentPath);
                }
                try
                {
                    var attachmentFilePath = temporaryAttachmentPath + "\\" + attachment.FileName;
                    attachment.SaveAsFile(attachmentFilePath);
                    result = System.IO.File.ReadAllBytes(attachmentFilePath);
                }
                catch (COMException ex)
                {
                    try
                    {
                        Log.Warn("Failed to get attachment bytes for " + attachment.DisplayName, ex);
                        // Swallow exception(!)

                        string strName = temporaryAttachmentPath + "\\" + DateTime.Now.ToString("MMddyyyyHHmmssfff") + ".html";
                        olItem.SaveAs(strName, Microsoft.Office.Interop.Outlook.OlSaveAsType.olHTML);
                        foreach (string strFileName in System.IO.Directory.GetFiles(strName.Replace(".html", "_files")))
                        {
                            if (strFileName.EndsWith("\\" + attachment.DisplayName))
                            {
                                result = System.IO.File.ReadAllBytes(strFileName);
                                break;
                            }
                        }
                    }
                    catch (Exception ex1)
                    {
                        Log.Error("EmailArchiving.GetAttachmentBytes", ex1);
                    }
                }
                finally
                {
                    if (System.IO.Directory.Exists(temporaryAttachmentPath))
                    {
                        System.IO.Directory.Delete(temporaryAttachmentPath, true);
                    }
                }
            }

            return result;
        }


        /// <summary>
        /// What date/time should be assigned to the archived email?
        /// </summary>
        /// <param name="olItem">The email to be archived.</param>
        /// <param name="reason">The reason the email is being archived.</param>
        /// <returns>An appropriate date time.</returns>
        public static DateTime ArchiveTime(this Outlook.MailItem olItem, EmailArchiveReason reason)
        {
            DateTime result;
            var now = DateTime.UtcNow;

            switch (reason)
            {
                case EmailArchiveReason.Outbound:
                case EmailArchiveReason.SendAndArchive:
                    result = olItem.CreationTime;
                    if (result > now)
                    {
                        /* if the actual date hasn't yet been set, Outlook will
                         * nonchalantly return 1st January 4501 */
                        result = now;
                    }
                    break;
                case EmailArchiveReason.Inbound:
                default:
                    result = olItem.SentOn;
                    break;
            }

            return result;
        }


        /// <summary>
        /// Archive this email item to CRM.
        /// </summary>
        /// <param name="olItem">The email item to archive.</param>
        /// <param name="reason">The reason it is being archived.</param>
        /// <returns>A result object indicating success or failure.</returns>
        public static ArchiveResult Archive(this Outlook.MailItem olItem, EmailArchiveReason reason, string excludedEmails = "")
        {
            ArchiveResult result;
            Outlook.UserProperty olProperty = olItem.UserProperties[CrmIdPropertyName];

            if (olProperty == null)
            {
                result = olItem.AsArchiveable(reason).Save(excludedEmails);
                
                if (result.IsSuccess)
                {
                    olItem.Categories = string.IsNullOrEmpty(olItem.Categories) ?
                        SuiteCRMCategoryName :
                        $"{olItem.Categories},{SuiteCRMCategoryName}";
                    olItem.EnsureProperty(CrmIdPropertyName, result.EmailId);
                }
            }
            else
            {
                result = ArchiveResult.Success(olProperty.Value, new[] { new AlreadyArchivedException(olItem) });
            }

            return result;
        }


        /// <summary>
        /// Ensure that I have a user property with this name and this value.
        /// </summary>
        /// <param name="olItem">Me</param>
        /// <param name="propertyName">The name of the property I should have.</param>
        /// <param name="propertyValue">The value of the property I should have.</param>
        private static void EnsureProperty(this Outlook.MailItem olItem, string propertyName, string propertyValue)
        {
            try
            {
                Outlook.UserProperty olProperty = olItem.UserProperties[propertyName];
                if (olProperty == null)
                {
                    olProperty = olItem.UserProperties.Add(propertyName, Outlook.OlUserPropertyType.olText);
                }
                olProperty.Value = propertyValue;
            }
            finally
            {
                olItem.Save();
            }
        }
    }
}
