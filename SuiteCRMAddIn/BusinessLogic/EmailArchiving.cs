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
namespace SuiteCRMAddIn.BusinessLogic
{
    using Daemon;
    using SuiteCRMClient;
    using SuiteCRMClient.Email;
    using SuiteCRMClient.Exceptions;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// The agent which handles the automatic and manual archiving of emails.
    /// </summary>
    /// <remarks>
    /// Some of functionality of this class is duplicated in SuiteCRMClient.Email.clsEmailArchive.
    /// TODO: Refactor. See issue #125
    /// </remarks>
    public class EmailArchiving : RepeatingProcess
    {
        private UserSession SuiteCRMUserSession => Globals.ThisAddIn.SuiteCRMUserSession;

        /// <summary>
        /// Magic property tag to get the email address from an Outlook Recipient object.
        /// </summary>
        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        /// <summary>
        /// Canonical format to use when saving date/times to CRM; essentially, ISO8601 without the 'T'.
        /// </summary>
        public const string EmailDateFormat = "yyyy-MM-dd HH:mm:ss";

        /// <summary>
        /// The name of the Outlook user property on which we will store the CRM Category associated
        /// with an email, of any.
        /// </summary>
        public const string CRMCategoryPropertyName = "SuiteCRMCategory";

        public EmailArchiving(string name, ILogger log) : base(name, log)
        {
        }

        internal override void PerformIteration()
        {
            if (Globals.ThisAddIn.HasCrmUserSession)
            {
                Log.Debug("Auto-Archive iteration started");

                var minReceivedDateTime = DateTime.UtcNow.AddDays(0 - Properties.Settings.Default.DaysOldEmailToAutoArchive);
                var foldersToBeArchived = GetMailFolders(Globals.ThisAddIn.Application.Session.Folders)
                    .Where(FolderShouldBeAutoArchived);

                foreach (var objFolder in foldersToBeArchived)
                {
                    ArchiveFolderItems(objFolder, minReceivedDateTime);
                }
                Log.Debug("Auto-Archive iteration completed");
            }
            else
            {
                Log.Debug("Auto-Archive iteration skipped because no user session.");
            }
        }

        private bool FolderShouldBeAutoArchived(Outlook.Folder folder) => FolderShouldBeAutoArchived(folder.EntryID);

        private bool FolderShouldBeAutoArchived(string folderEntryId)
            => Properties.Settings.Default.AutoArchiveFolders?.Contains(folderEntryId) ?? false;

        private void ArchiveFolderItems(Outlook.Folder objFolder, DateTime minReceivedDateTime)
        {
            try
            {
                var unreadEmails = objFolder.Items.Restrict(
                        $"[ReceivedTime] >= \'{minReceivedDateTime.AddDays(-1):yyyy-MM-dd HH:mm}\'");

                for (int intItr = 1; intItr <= unreadEmails.Count; intItr++)
                {
                    var objMail = unreadEmails[intItr] as Outlook.MailItem;
                    if (objMail != null)
                    {
                        // If this throws an exception here, we skip the rest of the folder
                        ArchiveNewMailItem(objMail, EmailArchiveType.Inbound);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error($"EmailArchiving.ArchiveFolderItems; folder {objFolder.Name}:", ex);
            }
        }

        public void ProcessEligibleNewMailItem(Outlook.MailItem objMail, EmailArchiveType archiveType)
        {
            var parentFolder = objMail.Parent as Outlook.Folder;
            if (parentFolder == null)
            {
                Log.Debug($"NULL email folder for {archiveType} “{objMail.Subject}”");
                return;
            }

            if (EmailShouldBeArchived(archiveType, parentFolder.Store))
            {
                ArchiveNewMailItem(objMail, archiveType);
            }
            else
            {
                Log.Debug($"NOT archiving {archiveType} email (folder {parentFolder.Name})");
            }
        }

        private bool EmailShouldBeArchived(EmailArchiveType type, Outlook.Store store)
        {
            var storeId = store.StoreID;
            switch (type)
            {
                case EmailArchiveType.Inbound:
                    return Properties.Settings.Default.AccountsToArchiveInbound.Contains(storeId);
                case EmailArchiveType.Sent:
                    return Properties.Settings.Default.AccountsToArchiveOutbound.Contains(storeId);
                default:
                    return false;
            }
        }

        public void ArchiveNewMailItem(Outlook.MailItem objMail, EmailArchiveType archiveType)
        {
            if (objMail.UserProperties["SuiteCRM"] == null)
            {
                bool archived = MaybeArchiveEmail(objMail, archiveType, Properties.Settings.Default.ExcludedEmails);
                objMail.UserProperties.Add("SuiteCRM", Outlook.OlUserPropertyType.olText, true, Outlook.OlUserPropertyType.olText);
                objMail.UserProperties["SuiteCRM"].Value = archived ? Boolean.TrueString : Boolean.FalseString;
                if (archived)
                {
                    objMail.Categories = "SuiteCRM";
                }
                objMail.Save();
            }
        }

        /// <summary>
        /// Get the item with this entry id.
        /// </summary>
        /// <param name="entryId">An outlook entry id.</param>
        /// <returns>the requested item, if found.</returns>
        public Outlook.MailItem GetItemById(string entryId)
        {
            return Globals.ThisAddIn.Application.GetNamespace("MAPI").GetItemFromID(entryId);
        }

        private bool MaybeArchiveEmail(Outlook.MailItem mailItem, EmailArchiveType archiveType, string strExcludedEmails = "")
        {
            bool result = false;
            var objEmail = SerialiseEmailObject(mailItem, archiveType);
            List<string> contacts = objEmail.GetValidContactIDs(strExcludedEmails);

            if (contacts.Count > 0)
            {
                Log.Info($"Archiving {archiveType} email “{mailItem.Subject}”");
                DaemonWorker.Instance.AddTask(new ArchiveEmailAction(SuiteCRMUserSession, objEmail, archiveType, contacts));
                result = true;
            }

            return result;
        }

        private clsEmailArchive SerialiseEmailObject(Outlook.MailItem mail, EmailArchiveType archiveType)
        {
            clsEmailArchive mailArchive = new clsEmailArchive(SuiteCRMUserSession, Log);
            mailArchive.From = ExtractSmtpAddressForSender(mail);
            mailArchive.To = string.Empty;

            Log.Info($"EmailArchiving.SerialiseEmailObject: serialising mail {mail.Subject} dated {mail.SentOn}.");

            foreach (Outlook.Recipient objRecepient in mail.Recipients)
            {
                string address = GetSmtpAddress(objRecepient);

                if (mailArchive.To == string.Empty)
                {
                    mailArchive.To = address;
                }
                else
                {
                    mailArchive.To += ";" + address;
                }
            }

            mailArchive.OutlookId = mail.EntryID;
            mailArchive.Subject = mail.Subject;
            mailArchive.Sent = DateTimeOfMailItem(mail, "autoOUTBOUND");
            mailArchive.Body = mail.Body;
            mailArchive.HTMLBody = mail.HTMLBody;
            mailArchive.ArchiveType = archiveType;
            if (Properties.Settings.Default.ArchiveAttachments)
            {
                foreach (Outlook.Attachment objMailAttachments in mail.Attachments)
                {
                    mailArchive.Attachments.Add(new clsEmailAttachments
                    {
                        DisplayName = objMailAttachments.DisplayName,
                        FileContentInBase64String = GetAttachmentBytes(objMailAttachments, mail)
                    });
                }
            }

            return mailArchive;
        }


        /// <summary>
        /// From this email recipient, extract the SMTP address (if that's possible).
        /// </summary>
        /// <param name="recipient">A recipient object</param>
        /// <returns>The SMTP address for that object, if it can be recovered, else an empty string.</returns>
        private string GetSmtpAddress(Outlook.Recipient recipient)
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
                    this.Log.Warn($"{this.GetType().Name}.ExtractSmtpAddressForSender: unknown email type {recipient.AddressEntry.Type}");
                    break;
            }

            return result;
        }

        /// <summary>
        /// From this mail item, extract the SMTP sender address if any, else the
        /// empty string.
        /// </summary>
        /// <remarks>
        /// If the sender is using Exchange (which if they're using Outlook they almost
        /// certainly are) the 'sender email address' won't be an email address, it will
        /// be a bizarre LDAP query which CRM will barf on. However, the Sender property
        /// may well be null, so allow for that too.
        /// </remarks>
        /// <param name="mail">The mail item</param>
        /// <returns>An SMTP address or an empty string.</returns>
        private string ExtractSmtpAddressForSender(Outlook.MailItem mail)
        {
            string result = string.Empty;

            try
            {
                switch (mail.SenderEmailType)
                {
                    case "SMTP":
                        result = mail.SenderEmailAddress;
                        break;
                    case "EX": /* an Exchange address */
                        var sender = mail.Sender;
                        if (sender != null)
                        {
                            var exchangeUser = sender.GetExchangeUser();
                            if (exchangeUser != null)
                            {
                                result = exchangeUser.PrimarySmtpAddress;
                            }
                        }

                        if (string.IsNullOrEmpty(result))
                        {
                            var currentUser = Globals.ThisAddIn.Application.ActiveExplorer().Session.CurrentUser.PropertyAccessor;
                            result = currentUser.GetProperty(PR_SMTP_ADDRESS).ToString();
                        }
                        break;
                    default:
                        this.Log.Warn($"{this.GetType().Name}.ExtractSmtpAddressForSender: unknown email type {mail.SenderEmailType}");
                        break;
                }
            }
            catch (Exception any)
            {
                this.Log.Error($"{this.GetType().Name}.ExtractSmtpAddressForSender: unexpected error", any);
            }

            return result;
        }

        private void ArchiveEmailThread(clsEmailArchive objEmail, EmailArchiveType archiveType, string strExcludedEmails = "")
        {
            try
            {
                if (SuiteCRMUserSession != null)
                {
                    while (SuiteCRMUserSession.AwaitingAuthentication == true)
                    {
                        Thread.Sleep(1000);
                    }
                    if (SuiteCRMUserSession.IsLoggedIn)
                    {
                        objEmail.SuiteCRMUserSession = SuiteCRMUserSession;
                        objEmail.Save(strExcludedEmails);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.ArchiveEmailThread", ex);
            }
        }

        public byte[] GetAttachmentBytes(Outlook.Attachment attachment, Outlook.MailItem mail)
        {
            byte[] strRet = null;

            Log.Info($"EmailArchiving.GetAttachmentBytes: serialising attachment '{attachment.FileName}' of email '{mail.Subject}'.");

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
                    strRet = System.IO.File.ReadAllBytes(attachmentFilePath);
                }
                catch (COMException ex)
                {
                    try
                    {
                        Log.Warn("Failed to get attachment bytes for " + attachment.DisplayName, ex);
                        // Swallow exception(!)

                        string strName = temporaryAttachmentPath + "\\" + DateTime.Now.ToString("MMddyyyyHHmmssfff") + ".html";
                        mail.SaveAs(strName, Microsoft.Office.Interop.Outlook.OlSaveAsType.olHTML);
                        foreach (string strFileName in System.IO.Directory.GetFiles(strName.Replace(".html", "_files")))
                        {
                            if (strFileName.EndsWith("\\" + attachment.DisplayName))
                            {
                                strRet = System.IO.File.ReadAllBytes(strFileName);
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

            return strRet;
        }

        private IEnumerable<Outlook.Folder> GetMailFolders(Outlook.Folders root)
        {
            var result = new List<Outlook.Folder>();
            GetMailFoldersHelper(root, result);
            return result;
        }

        private void GetMailFoldersHelper(Outlook.Folders objInpFolders, IList<Outlook.Folder> result)
        {
            try
            {
                foreach (Outlook.Folder objFolder in objInpFolders)
                {
                    if (objFolder.Folders.Count > 0)
                    {
                        result.Add(objFolder);
                        GetMailFoldersHelper(objFolder.Folders, result);
                    }
                    else
                        result.Add(objFolder);
                }
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.GetMailFolders", ex);
                ;
            }
        }

        public ArchiveResult ArchiveEmailWithEntityRelationships(Outlook.MailItem mailItem, IEnumerable<CrmEntity> selectedCrmEntities, string type)
        {
            var result = this.SaveEmailToCrm(mailItem, type);
            if (result.IsSuccess)
            {
                var warnings = CreateEmailRelationshipsWithEntities(result.EmailId, selectedCrmEntities);
                result = ArchiveResult.Success(
                    result.EmailId,
                    result.Problems == null ?
                    warnings :
                    result.Problems.Concat(warnings));
            }

            return result;
        }

        private IList<System.Exception> CreateEmailRelationshipsWithEntities(string crmMailId, IEnumerable<CrmEntity> selectedCrmEntities)
        {
            var failures = new List<System.Exception>();
            foreach (var entity in selectedCrmEntities)
            {
                try
                {
                    CreateEmailRelationshipOrFail(crmMailId, entity);
                }
                catch (System.Exception failure)
                {
                    Log.Error("CreateEmailRelationshipsWithEntities", failure);
                    failures.Add(failure);
                }
            }
            return failures;
        }

        private void SaveMailItemIfNecessary(Outlook.MailItem o, string type)
        {
            if (type == "SendArchive")
            {
                o.Save();
            }
        }

        /// <summary>
        /// Save this email item, of this type, to CRM.
        /// </summary>
        /// <param name="mailItem">The email item to send.</param>
        /// <param name="type">?unknown</param>
        /// <returns>An archive result comprising the CRM id of the email, if stored,
        /// and a list of exceptions encountered in the process.</returns>
        public ArchiveResult SaveEmailToCrm(Outlook.MailItem mailItem, string type)
        {
            ArchiveResult result;
            try
            {
                SaveMailItemIfNecessary(mailItem, type);

                result = ConstructAndDespatchCrmItem(mailItem, type);

                if (!String.IsNullOrEmpty(result.EmailId))
                {
                    /* we successfully saved the email item itself */
                    mailItem.Categories = "SuiteCRM";
                    mailItem.Save();

                    if (mailItem.Attachments.Count > 0)
                    {
                        result = ConstructAndDespatchAttachments(mailItem, result);
                    }
                }
            }
            catch (System.Exception failure)
            {
                Log.Warn("Could not upload email to CRM", failure);
                result = ArchiveResult.Failure(failure);
            }

            return result;
        }

        /// <summary>
        /// Construct and despatch CRM representations of the attachments of this email item to CRM.
        /// </summary>
        /// <param name="mailItem">The mail item whose attachments should be sent.</param>
        /// <param name="result">The result of transmitting the item itself to CRM.</param>
        /// <returns>A (possibly modified) archive result.</returns>
        private ArchiveResult ConstructAndDespatchAttachments(Outlook.MailItem mailItem, ArchiveResult result)
        {
            var warnings = new List<System.Exception>();

            if (Properties.Settings.Default.ArchiveAttachments)
            {
                foreach (Outlook.Attachment attachment in mailItem.Attachments)
                {
                    warnings.Add(ConstructAndDespatchCrmAttachment(mailItem, result.EmailId, attachment));
                }
            }

            if (warnings.Where(w => w != null).Count() > 0)
            {
                if (result.Problems != null)
                {
                    warnings.AddRange(result.Problems);
                }
                result = ArchiveResult.Success(result.EmailId, warnings.Where(w => w != null));
            }

            return result;
        }

        /// <summary>
        /// Construct and despatch a CRM representation of this attachment to CRM.
        /// </summary>
        /// <param name="mailItem">The mail item to which this attachment is attached.</param>
        /// <param name="crmId">The id of that mail item in CRM.</param>
        /// <param name="attachment">The attachment to despatch.</param>
        /// <returns>Any exception which was thrown while attempting to despatch the attachment.</returns>
        private Exception ConstructAndDespatchCrmAttachment(Outlook.MailItem mailItem, string crmId, Outlook.Attachment attachment)
        {
            Exception result = null;
            try
            {
                clsSuiteCRMHelper.UploadAttachment(
                    new clsEmailAttachments
                    {
                        DisplayName = attachment.DisplayName,
                        FileContentInBase64String = GetAttachmentBytes(attachment, mailItem)
                    },
                    crmId);
            }
            catch (System.Exception problem)
            {
                Log.Warn("Failed to upload email attachment", problem);
                result = problem;
            }

            return result;
        }

        /// <summary>
        /// Construct and despatch a CRM representation of this mail item, without its attachments, to CRM
        /// </summary>
        /// <param name="mailItem">The mail item to despatch.</param>
        /// <param name="type">?unknown.</param>
        /// <returns>An archive result comprising the CRM id of the email, if stored,
        /// and a list of exceptions encountered in the process.</returns>
        private ArchiveResult ConstructAndDespatchCrmItem(Outlook.MailItem mailItem, string type)
        {
            ArchiveResult result;
            eNameValue[] crmItem = ConstructCrmItem(mailItem, type);

            try
            {
                result = ArchiveResult.Success(clsSuiteCRMHelper.SetEntry(crmItem, "Emails"), null);
            }
            catch (System.Exception firstFailure)
            {
                Log.Warn("EmailArchiving.SaveEmailToCrm: first attempt to upload email failed", firstFailure);

                try
                {
                    /* try again without the HTML body. I have no idea why this might make a difference. */
                    crmItem[5] = clsSuiteCRMHelper.SetNameValuePair("description_html", string.Empty);

                    result = ArchiveResult.Success(clsSuiteCRMHelper.SetEntry(crmItem, "Emails"), new[] { firstFailure });
                }
                catch (System.Exception secondFailure)
                {
                    Log.Warn("EmailArchiving.SaveEmailToCrm: second attempt to upload email (without HTML body) failed", firstFailure);
                    result = ArchiveResult.Failure(new[] { firstFailure, secondFailure });
                }
            }

            return result;
        }

        /// <summary>
        /// Construct a CRM representation of this mail item, without its attachments if any.
        /// </summary>
        /// <param name="mailItem">The mail item.</param>
        /// <param name="type">?unknown.</param>
        /// <returns>A CRM representation of the item, as a set of name/value pairs.</returns>
        private eNameValue[] ConstructCrmItem(Outlook.MailItem mailItem, string type)
        {
            eNameValue[] data = new eNameValue[13];
            string category = mailItem.UserProperties[CRMCategoryPropertyName] != null ?
                mailItem.UserProperties[CRMCategoryPropertyName].Value :
                string.Empty;

            data[0] = clsSuiteCRMHelper.SetNameValuePair("name", mailItem.Subject ?? string.Empty);
            data[1] = clsSuiteCRMHelper.SetNameValuePair("date_sent", DateTimeOfMailItem(mailItem, type).ToString(EmailDateFormat));
            data[2] = clsSuiteCRMHelper.SetNameValuePair("message_id", mailItem.EntryID);
            data[3] = clsSuiteCRMHelper.SetNameValuePair("status", "archived");
            data[4] = clsSuiteCRMHelper.SetNameValuePair("description", mailItem.Body ?? string.Empty);
            data[5] = clsSuiteCRMHelper.SetNameValuePair("description_html", mailItem.HTMLBody ?? string.Empty);
            data[6] = clsSuiteCRMHelper.SetNameValuePair("from_addr", clsGlobals.GetSenderAddress(mailItem, type));
            data[7] = clsSuiteCRMHelper.SetNameValuePair("to_addrs", mailItem.To);
            data[8] = clsSuiteCRMHelper.SetNameValuePair("cc_addrs", mailItem.CC);
            data[9] = clsSuiteCRMHelper.SetNameValuePair("bcc_addrs", mailItem.BCC);
            data[10] = clsSuiteCRMHelper.SetNameValuePair("reply_to_addr", mailItem.ReplyRecipientNames);
            data[11] = clsSuiteCRMHelper.SetNameValuePair("assigned_user_id", clsSuiteCRMHelper.GetUserId());
            data[12] = clsSuiteCRMHelper.SetNameValuePair("category_id", category);
            return data;
        }

        private DateTime DateTimeOfMailItem(Outlook.MailItem mailItem, string type)
        {
            DateTime result;
            var now = DateTime.UtcNow;

            switch (type)
            {
                case "autoOUTBOUND":
                    result = mailItem.CreationTime;
                    if (result > now)
                    {
                        /* if the actual date hasn't yet been set, Outlook will
                         * nonchalantly return 1st January 4501 */
                        result = now;
                    }
                    break;
                case "SendArchive":
                    result = mailItem.CreationTime;
                    break;
                case null:
                case "autoINBOUND":
                default:
                    result = mailItem.SentOn;
                    break;
            }
            return result;
        }

        public void CreateEmailRelationshipOrFail(string emailId, CrmEntity entity)
        {
            var success = clsSuiteCRMHelper.TrySetRelationship(
                new eSetRelationshipValue
                {
                    module2 = "emails",
                    module2_id = emailId,
                    module1 = entity.ModuleName,
                    module1_id = entity.EntityId,
                }, Objective.Email);

            if (!success) throw new CrmSaveDataException($"Cannot create email relationship with {entity.ModuleName} ('set_relationship' failed)");
        }
    }
}
