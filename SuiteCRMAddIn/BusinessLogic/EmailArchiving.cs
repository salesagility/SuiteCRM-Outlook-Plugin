using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Runtime.InteropServices;
using SuiteCRMClient;
using SuiteCRMClient.Email;
using SuiteCRMClient.Exceptions;
using SuiteCRMClient.Logging;
using SuiteCRMClient.RESTObjects;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    public class EmailArchiving
    {
        private ILogger Log => Globals.ThisAddIn.Log;

        private clsUsersession SuiteCRMUserSession => Globals.ThisAddIn.SuiteCRMUserSession;

        private clsSettings settings => Globals.ThisAddIn.settings;

        public void ArchiveMailInAutoArchiveFolders()
        {
            Log.Debug("Auto-Archive thread started");
            var minReceivedDateTime = DateTime.UtcNow.AddDays(0 - settings.DaysOldEmailToAutoArchive);
            try
            {
                var foldersToBeArchived = GetMailFolders(Globals.ThisAddIn.Application.Session.Folders)
                    .Where(FolderShouldBeAutoArchived);

                foreach (var objFolder in foldersToBeArchived)
                {
                    ArchiveFolderItems(objFolder, minReceivedDateTime);
                }
            }
            catch (Exception ex)
            {
                Log.Error("ArchiveMailInAutoArchiveFolders", ex);
            }
            finally
            {
                Log.Info("Auto-Archive thread completed");
            }
        }

        private bool FolderShouldBeAutoArchived(Outlook.Folder folder) => FolderShouldBeAutoArchived(folder.EntryID);

        private bool FolderShouldBeAutoArchived(string folderEntryId)
            => settings.AutoArchiveFolders?.Contains(folderEntryId) ?? false;

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
                Log.Error("EmailArchiving.ProcessFolderItems", ex);
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
                    return settings.AccountsToArchiveInbound.Contains(storeId);
                case EmailArchiveType.Sent:
                    return settings.AccountsToArchiveOutbound.Contains(storeId);
                default:
                    return false;
            }
        }

        public void ArchiveNewMailItem(Outlook.MailItem objMail, EmailArchiveType archiveType)
        {
            if (objMail.UserProperties["SuiteCRM"] == null)
            {
                ArchiveEmail(objMail, archiveType, this.settings.ExcludedEmails);
                objMail.UserProperties.Add("SuiteCRM", Outlook.OlUserPropertyType.olText, true, Outlook.OlUserPropertyType.olText);
                objMail.UserProperties["SuiteCRM"].Value = "True";
                objMail.Categories = "SuiteCRM";
                objMail.Save();
            }
        }

        private void ArchiveEmail(Outlook.MailItem objMail, EmailArchiveType archiveType, string strExcludedEmails = "")
        {
            Log.Info($"Archiving {archiveType} email “{objMail.Subject}”");
            var objEmail = SerialiseEmailObject(objMail, archiveType);
            Thread objThread = new Thread(() => ArchiveEmailThread(objEmail, archiveType, strExcludedEmails));
            objThread.Start();
        }

        private clsEmailArchive SerialiseEmailObject(Outlook.MailItem objMail, EmailArchiveType archiveType)
        {
            clsEmailArchive objEmail = new clsEmailArchive();
            objEmail.From = objMail.SenderEmailAddress;
            objEmail.To = "";
            foreach (Outlook.Recipient objRecepient in objMail.Recipients)
            {
                if (objEmail.To == "")
                    objEmail.To = objRecepient.Address;
                else
                    objEmail.To += ";" + objRecepient.Address;
            }
            objEmail.Subject = objMail.Subject;
            objEmail.Body = objMail.Body;
            objEmail.HTMLBody = objMail.HTMLBody;
            objEmail.ArchiveType = archiveType;
            foreach (Outlook.Attachment objMailAttachments in objMail.Attachments)
            {
                objEmail.Attachments.Add(new clsEmailAttachments
                {
                    DisplayName = objMailAttachments.DisplayName,
                    FileContentInBase64String = GetAttachmentBytes(objMailAttachments, objMail)
                });
            }
            return objEmail;
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

        public byte[] GetAttachmentBytes(Outlook.Attachment objMailAttachment, Outlook.MailItem objMail)
        {
            byte[] strRet = null;
            if (objMailAttachment != null)
            {
                var temporaryAttachmentPath = Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath";
                if (!System.IO.Directory.Exists(temporaryAttachmentPath))
                {
                    System.IO.Directory.CreateDirectory(temporaryAttachmentPath);
                }
                try
                {
                    var attachmentFilePath = temporaryAttachmentPath + "\\" + objMailAttachment.FileName;
                    objMailAttachment.SaveAsFile(attachmentFilePath);
                    strRet = System.IO.File.ReadAllBytes(attachmentFilePath);
                }
                catch (COMException ex)
                {
                    try
                    {
                        Log.Warn("Failed to get attachment bytes for " + objMailAttachment.DisplayName, ex);
                        // Swallow exception(!)

                        string strName = temporaryAttachmentPath + "\\" + DateTime.Now.ToString("MMddyyyyHHmmssfff") + ".html";
                        objMail.SaveAs(strName, Microsoft.Office.Interop.Outlook.OlSaveAsType.olHTML);
                        foreach (string strFileName in System.IO.Directory.GetFiles(strName.Replace(".html", "_files")))
                        {
                            if (strFileName.EndsWith("\\" + objMailAttachment.DisplayName))
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

        public ArchiveResult ArchiveEmailWithEntityRelationships(Outlook.MailItem mailItem, List<CrmEntity> selectedCrmEntities, string type)
        {
            var result = this.SaveEmailToCrm(mailItem, type);
            if (result.IsFailure) return result;
            var warnings = CreateEmailRelationshipsWithEntities(result.EmailId, selectedCrmEntities);
            return ArchiveResult.Success(
                result.EmailId,
                result.Problems.Concat(warnings));
        }

        private IList<System.Exception> CreateEmailRelationshipsWithEntities(string crmMailId, List<CrmEntity> selectedCrmEntities)
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

        public ArchiveResult SaveEmailToCrm(Outlook.MailItem mailItem, string type)
        {
            try
            {
                SaveMailItemIfNecessary(mailItem, type);

                eNameValue[] data = new eNameValue[12];
                data[0] = clsSuiteCRMHelper.SetNameValuePair("name", mailItem.Subject ?? "");
                data[1] = clsSuiteCRMHelper.SetNameValuePair("date_sent", DateTimeOfMailItem(mailItem, type).ToString("yyyy-MM-dd HH:mm:ss"));
                data[2] = clsSuiteCRMHelper.SetNameValuePair("message_id", mailItem.EntryID);
                data[3] = clsSuiteCRMHelper.SetNameValuePair("status", "archived");
                data[4] = clsSuiteCRMHelper.SetNameValuePair("description", mailItem.Body ?? "");
                data[5] = clsSuiteCRMHelper.SetNameValuePair("description_html", mailItem.HTMLBody);
                data[6] = clsSuiteCRMHelper.SetNameValuePair("from_addr", clsGlobals.GetSenderAddress(mailItem, type));
                data[7] = clsSuiteCRMHelper.SetNameValuePair("to_addrs", mailItem.To);
                data[8] = clsSuiteCRMHelper.SetNameValuePair("cc_addrs", mailItem.CC);
                data[9] = clsSuiteCRMHelper.SetNameValuePair("bcc_addrs", mailItem.BCC);
                data[10] = clsSuiteCRMHelper.SetNameValuePair("reply_to_addr", mailItem.ReplyRecipientNames);
                data[11] = clsSuiteCRMHelper.SetNameValuePair("assigned_user_id", clsSuiteCRMHelper.GetUserId());

                string crmEmailId;
                try
                {
                    crmEmailId = clsSuiteCRMHelper.SetEntry(data, "Emails");
                }
                catch (System.Exception firstFailure)
                {
                    Log.Warn("1st attempt to upload email failed", firstFailure);
                    data[5] = clsSuiteCRMHelper.SetNameValuePair("description_html", "");
                    try
                    {
                        crmEmailId = clsSuiteCRMHelper.SetEntry(data, "Emails");
                    }
                    catch (System.Exception secondFailure)
                    {
                        Log.Warn("2nd attempt to upload email failed", secondFailure);
                        return ArchiveResult.Failure(new[] { firstFailure, secondFailure });
                    }
                }

                mailItem.Categories = "SuiteCRM";
                mailItem.Save();
                var warnings = new List<System.Exception>();
                if (settings.ArchiveAttachmentsDefault)
                {
                    foreach (Outlook.Attachment attachment in mailItem.Attachments)
                    {
                        try
                        {
                            clsSuiteCRMHelper.UploadAttachment(
                                new clsEmailAttachments
                                {
                                    DisplayName = attachment.DisplayName,
                                    FileContentInBase64String = GetAttachmentBytes(attachment, mailItem)
                                },
                                crmEmailId);
                        }
                        catch (System.Exception problem)
                        {
                            Log.Warn("Failed to upload email attachment", problem);
                            warnings.Add(problem);
                        }
                    }
                }
                return ArchiveResult.Success(crmEmailId, warnings);
            }
            catch (System.Exception failure)
            {
                Log.Warn("Could not upload email to CRM", failure);
                return ArchiveResult.Failure(failure);
            }
        }

        private DateTime DateTimeOfMailItem(Outlook.MailItem mailItem, string type)
        {
            DateTime dateTime;
            switch (type)
            {
                case "autoOUTBOUND":
                case "SendArchive":
                    dateTime = mailItem.CreationTime;
                    break;
                case null:
                case "autoINBOUND":
                default:
                    dateTime = mailItem.SentOn;
                    break;
            }
            return dateTime;
        }

        public void CreateEmailRelationshipOrFail(string emailId, CrmEntity entity)
        {
            var success = clsSuiteCRMHelper.SetRelationship(
                new eSetRelationshipValue
                {
                    module2 = "emails",
                    module2_id = emailId,
                    module1 = entity.ModuleName,
                    module1_id = entity.EntityId,
                });

            if (!success) throw new CrmSaveDataException($"Cannot create email relationship with {entity.ModuleName} ('set_relationship' failed)");
        }
    }
}
