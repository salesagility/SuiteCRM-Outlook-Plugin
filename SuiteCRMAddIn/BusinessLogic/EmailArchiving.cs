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
    using SuiteCRMClient;
    using SuiteCRMClient.Email;
    using SuiteCRMClient.Exceptions;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using SuiteCRMAddIn.Extensions;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    /// <summary>
    /// The agent which handles the automatic and manual archiving of emails.
    /// </summary>
    /// <remarks>
    /// Some of functionality of this class is duplicated in SuiteCRMClient.Email.clsEmailArchive.
    /// TODO: Refactor. See issue #125
    /// </remarks>
    public class EmailArchiving : RepeatingProcess
    {
        /// <summary>
        /// Convenience property to get a handle on the global user session.
        /// </summary>
        private UserSession SuiteCRMUserSession => Globals.ThisAddIn.SuiteCRMUserSession;

        /// <summary>
        /// Canonical format to use when saving date/times to CRM; essentially, ISO8601 without the 'T'.
        /// </summary>
        public const string EmailDateFormat = "yyyy-MM-dd HH:mm:ss";

        /// <summary>
        /// The modules to which we'll try to save if no more specific list of modules is specified.
        /// </summary>
        public static readonly List<string> defaultModuleKeys = new List<string>() { ContactSynchroniser.CrmModule, "Leads", "Accounts" };

        public EmailArchiving(string name, ILogger log) : base(name, log)
        {
        }

        internal override void PerformIteration()
        {
            if (Globals.ThisAddIn.HasCrmUserSession)
            {
                var minReceivedDateTime = DateTime.UtcNow.AddDays(0 - Properties.Settings.Default.DaysOldEmailToAutoArchive);
                var foldersToBeArchived = GetMailFolders(Globals.ThisAddIn.Application.Session.Folders)
                    .Where(FolderShouldBeAutoArchived);

                if (foldersToBeArchived.Count() > 0)
                {
                    Log.Debug("Auto-Archive iteration started");

                    foreach (var folder in foldersToBeArchived)
                    {
                        ArchiveFolderItems(folder, minReceivedDateTime);
                    }
                    Log.Debug("Auto-Archive iteration completed");
                }
                else
                {
                    Log.Debug("No folders to auto-archive.");
                }
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
            this.ArchiveFolderItems(objFolder, minReceivedDateTime, defaultModuleKeys);
        }


        /// <summary>
        /// Archive items in the specified folder which are email items, and which have been 
        /// received since the specified date.
        /// </summary>
        /// <remarks>
        /// I don't understand all of this. I particularly don't understand why we ever call it 
        /// on folders whose content are not mail items.
        /// </remarks>
        /// <param name="folder">The folder to archive.</param>
        /// <param name="minReceivedDateTime">The date to search from.</param>
        /// <param name="moduleKeys">The keys of the modules to which we'll seek to relate the archived item.</param>
        private void ArchiveFolderItems(Outlook.Folder folder, DateTime minReceivedDateTime, IEnumerable<string> moduleKeys)
        {
            try
            {
                /* safe but undesirable fallback - if we cannot identify a property to restrict by,
                 * search the whole folder */
                Outlook.Items candidateItems = folder.Items;

                if (folder.DefaultItemType == Outlook.OlItemType.olMailItem)
                {
                    foreach (string property in new string[] { "ReceivedTime", "LastModificationTime" })
                    {
                        try
                        {
                            candidateItems = folder.Items.Restrict(
                              $"[{property}] >= \'{minReceivedDateTime.AddDays(-1):yyyy-MM-dd HH:mm}\'");
                            break;
                        }
                        catch (COMException)
                        {
                            Log.Warn($"EmailArchiving.ArchiveFolderItems; Items in folder {folder.Name} do not have a {property} property");
                        }
                    }

                    foreach (var candidate in candidateItems)
                    {
                        var comType = Microsoft.VisualBasic.Information.TypeName(candidate);
                        
                        switch (comType)
                        {
                            case "MailItem":
                                ArchiveMailItem(candidate, moduleKeys);
                                break;
                            case "MeetingItem":
                            case "ReportItem":
                                Log.Debug($"EmailArchiving.ArchiveFolderItems; candidate is a '{comType}', we don't archive these");
                                break;
                            default:
                                Log.Debug($"EmailArchiving.ArchiveFolderItems; candidate is a '{comType}', don't know how to archive these");
                                break;
                        }
                    }
                }
                else
                {
                    Log.Debug($"EmailArchiving.ArchiveFolderItems; Folder {folder.Name} does not contain mail items, not archiving");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle($"Failed while archiving and email item in folder {folder.Name}:", ex);
            }
        }


        /// <summary>
        /// Archive an item believed to be an Outlook.MailItem.
        /// </summary>
        /// <param name="item">The item to archive.</param>
        /// <param name="moduleKeys">Keys of module(s) to relate the item to.</param>
        private void ArchiveMailItem(object item, IEnumerable<string> moduleKeys)
        {
            var olItem = item as Outlook.MailItem;
            if (olItem != null)
            {
                try
                {
                    olItem.Archive(EmailArchiveReason.Inbound, moduleKeys.Select(x => new CrmEntity(x, null)));
                }
                catch (Exception any)
                {
                    ErrorHandler.Handle($"Failed to archive MailItem '{olItem.Subject}' from '{olItem.GetSenderSMTPAddress()}", any);
                }
            }
        }


        public void ProcessEligibleNewMailItem(Outlook.MailItem olItem, EmailArchiveReason reason, string excludedEmails = "")
        {
            var parentFolder = olItem.Parent as Outlook.Folder;
            if (parentFolder == null)
            {
                Log.Debug($"NULL email folder for {reason} “{olItem.Subject}”");
                return;
            }

            if (EmailShouldBeArchived(reason, parentFolder.Store))
            {
                olItem.Archive(reason, defaultModuleKeys.Select(x => new CrmEntity(x, null)), excludedEmails);
            }
            else
            {
                Log.Debug($"NOT archiving {reason} email (folder {parentFolder.Name})");
            }
        }

        private bool EmailShouldBeArchived(EmailArchiveReason type, Outlook.Store store)
        {
            bool result;
            var storeId = store.StoreID;
            switch (type)
            {
                case EmailArchiveReason.Inbound:
                    result = Properties.Settings.Default.AccountsToArchiveInbound != null &&
                        Properties.Settings.Default.AccountsToArchiveInbound.Contains(storeId);
                    break;
                case EmailArchiveReason.Outbound:
                    result = Properties.Settings.Default.AccountsToArchiveOutbound != null &&
                        Properties.Settings.Default.AccountsToArchiveOutbound.Contains(storeId);
                    break;
                default:
                    result = false;
                    break;
            }

            return result;
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
                        try
                        {
                            result.Add(objFolder);
                            GetMailFoldersHelper(objFolder.Folders, result);
                        }
                        catch (COMException comx)
                        {
                            MessageBox.Show($"Failed to open mail folder {objFolder.Description} because {comx.Message}", "Failed to open mail folder", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            throw;
                        }
                    }
                    else
                    {
                        result.Add(objFolder);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle("Failed while trying to get mail folders", ex);
            }
        }


        public ArchiveResult ArchiveEmailWithEntityRelationships(Outlook.MailItem olItem, IEnumerable<CrmEntity> selectedCrmEntities, EmailArchiveReason reason)
        {
            return olItem.Archive(reason, selectedCrmEntities);
        }

        private void SaveMailItemIfNecessary(Outlook.MailItem olItem, EmailArchiveReason reason)
        {
            if (reason == EmailArchiveReason.SendAndArchive)
            {
                olItem.Save();
            }
        }


        public void CreateEmailRelationshipOrFail(CrmId emailId, CrmEntity entity)
        {
            var success = RestAPIWrapper.TrySetRelationship(
                new SetRelationshipParams
                {
                    module2 = "emails",
                    module2_id = emailId.ToString(),
                    module1 = entity.ModuleName,
                    module1_id = entity.EntityId,
                }, Objective.Email);

            if (!success) throw new CrmSaveDataException($"Cannot create email relationship with {entity.ModuleName} ('set_relationship' failed)");
        }
    }
}
