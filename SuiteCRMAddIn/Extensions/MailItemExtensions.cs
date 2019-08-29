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
    using Exceptions;
    using SuiteCRMClient;
    using SuiteCRMClient.Email;
    using SuiteCRMClient.Logging;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using TidyManaged;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Extension methods for Outlook MailItem objects.
    /// </summary>
    public static class MailItemExtensions
    {
        /// <summary>
        /// A cache of SMTP addresses, so we're not continually fetching them from a remote 
        /// Exchange server.
        /// </summary>
        private static Dictionary<Outlook.MailItem, string> senderAddressCache = new Dictionary<Outlook.MailItem, string>();

        /// <summary>
        /// Magic property tag to get the email address from an Outlook Recipient object.
        /// </summary>
        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        /// <summary>
        /// The name of the magic category we set when a mail is successfully archived.
        /// </summary>
        public const string SuiteCRMCategoryName = "SuiteCRM";

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

            if (senderAddressCache.ContainsKey(olItem))
            {
                result = senderAddressCache[olItem];
            }

            if (string.IsNullOrEmpty(result))
            {
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
                    ErrorHandler.Handle("Failed while trying to get the sender's SMTP address", any);
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
                    ErrorHandler.Handle("Failed to get email address of current user",
                        any);
                }

                senderAddressCache[olItem] = result;
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
            mailArchive.CrmEntryId = olItem.GetCRMEntryId();

            Log.Info($"MailItemExtension.AsArchiveable: serialising mail {olItem.Subject} dated {olItem.SentOn}.");

            foreach (Outlook.Recipient recipient in olItem.Recipients)
            {
                string address = recipient.GetSmtpAddress();

                switch (recipient.Type)
                {
                    case (int)Outlook.OlMailRecipientType.olCC:
                        mailArchive.CC = ExtendRecipientField(mailArchive.CC, address);
                        break;
                    case (int)Outlook.OlMailRecipientType.olBCC:
                        // unlikely to happen and in any case we don't store these
                        break;
                    default:
                        mailArchive.To = ExtendRecipientField(mailArchive.To, address);
                        break;
                }
            }

            mailArchive.CC = olItem.CC;

            mailArchive.ClientId = olItem.EnsureEntryID();
            mailArchive.Subject = olItem.Subject;
            mailArchive.Sent = olItem.ArchiveTime(reason);
            mailArchive.Body = olItem.Body;
            mailArchive.HTMLBody = Tidy(olItem.HTMLBody);
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

        private static string ExtendRecipientField(string fieldContent, string address)
        {
            return string.IsNullOrEmpty(fieldContent) ? address : $"{fieldContent};{address}";
        }


        /// <summary>
        /// The "HTML" which Outlook generates is diabolically bad, and CMS frequently chokes on it.
        /// Convert it to valid HTML before dispatch.
        /// </summary>
        /// <param name="html">The HTML - possibly including vile Microsoft junk - to tidy.</param>
        /// <returns>Nice clean XHTML.</returns>
        private static string Tidy(string html)
        {
            using (Document doc = Document.FromString(html))
            {
                doc.ShowWarnings = false;
                doc.Quiet = true;
                doc.OutputXhtml = true;
                doc.CleanAndRepair();
                return doc.Save();
            }
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
                        ErrorHandler.Handle("Failed to get data of an email attachment", ex1);
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

        public static ArchiveResult Archive(this Outlook.MailItem olItem, EmailArchiveReason reason)
        {
            return Archive(olItem, reason, EmailArchiving.defaultModuleKeys.Select(x => new CrmEntity(x, null)));
        }

        public static string GetCRMEntryId(this Outlook.MailItem olItem)
        {
            string result;
            Outlook.UserProperty property = null;
            
            try
            {
                property = olItem.UserProperties[SyncStateManager.CrmIdPropertyName];

                if (property == null)
                {
                    /* #6661: fail over to legacy property name if current property 
                     * name not found */
                    property = olItem.UserProperties[SyncStateManager.LegacyCrmIdPropertyName];
                }

                result = property != null ? property.Value.ToString() : string.Empty;
            }
            catch (COMException cex)
            {
                ErrorHandler.Handle("Could not get property while archiving email", cex);
                result = string.Empty;
            }

            return result;
        }

        /// <summary>
        /// Archive this email item to CRM.
        /// </summary>
        /// <param name="olItem">The email item to archive.</param>
        /// <param name="reason">The reason it is being archived.</param>
        /// <param name="moduleKeys">Keys (standardised names) of modules to search.</param>
        /// <param name="excludedEmails">email address(es) which should not be linked.</param>
        /// <returns>A result object indicating success or failure.</returns>
        public static ArchiveResult Archive(this Outlook.MailItem olItem, EmailArchiveReason reason, IEnumerable<CrmEntity> moduleKeys, string excludedEmails = "")
        {
            ArchiveResult result = olItem.AsArchiveable(reason).Save(moduleKeys, excludedEmails);

            if (result.IsSuccess)
            {
                try
                {
                    if (string.IsNullOrEmpty(olItem.Categories))
                    {
                        olItem.Categories = SuiteCRMCategoryName;
                    }
                    else if (olItem.Categories.IndexOf(SuiteCRMCategoryName) == -1)
                    {
                        olItem.Categories = $"{olItem.Categories},{SuiteCRMCategoryName}";
                    }

                    olItem.EnsureProperty(SyncStateManager.CrmIdPropertyName, result.EmailId);
                }
                catch (COMException cex)
                {
                    ErrorHandler.Handle("Could not set property while archiving email", cex);
                }
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

        /// <summary>
        /// Ensure that I have a user property with this name and this value.
        /// </summary>
        /// <param name="olItem">Me</param>
        /// <param name="propertyName">The name of the property I should have.</param>
        /// <param name="propertyValue">The value of the property I should have.</param>
        private static void EnsureProperty(this Outlook.MailItem olItem, string propertyName, object propertyValue)
        {
            EnsureProperty(olItem, propertyName, propertyValue.ToString());
        }
    }
}
