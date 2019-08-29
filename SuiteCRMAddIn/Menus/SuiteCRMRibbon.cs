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
    using Microsoft.Office.Core;
    using stdole;
    using SuiteCRMAddIn.BusinessLogic;
    using SuiteCRMAddIn.Dialogs;
    using SuiteCRMAddIn.Extensions;
    using SuiteCRMAddIn.Properties;
    using SuiteCRMClient.Email;
    using SuiteCRMClient.Logging;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using Office = Microsoft.Office.Core;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Code for our user interfave (ribbon, context menus) items. See also 
    /// MailRead.xml and MailRead2007.xml in this directory.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Note that the names `MailRead.xml` and `MailRead2007.xml` appear to be 
    /// magic; the files do not deal with the mail reading interface 
    /// exclusively, if they are renamed Outlook disables the addin.
    /// </para>
    /// <para>
    /// For more information, see the Ribbon XML documentation in the Visual 
    /// Studio Tools for Office Help.
    /// </para>
    /// </remarks>
    [ComVisible(true)]
    public class SuiteCRMRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        
        public SuiteCRMRibbon()
        {
        }

        public bool SuiteCRMMainTab_GetVisible(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer() != null)
            {
                if (Globals.ThisAddIn.Application.ActiveExplorer().Selection is Outlook.MailItem)
                {
                    return true;
                }
            }
            return false;
        }

        public IPictureDisp GetImage(IRibbonControl control)
        {
            IPictureDisp result;

            switch (control.Id)
            {
                case "btnSendAndArchive":
                    result = RibbonImageHelper.Convert(Resources.SendAndArchive);
                    break;
                case "btnSettings":
                    result = RibbonImageHelper.Convert(Resources.Settings);
                    break;
                case "btnAddressBook":
                    result = RibbonImageHelper.Convert(Resources.AddressBook);
                    break;
                case "manualSyncButton":
                case "manualSyncMultiButton":
                case "manualSyncToolbar":
                    result = RibbonImageHelper.Convert(Resources.manualSyncContact);
                    break;
                default:
                    result = RibbonImageHelper.Convert(Resources.Archive);
                    break;                
            }

            return result;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string result;

            switch (ribbonID)
            {
                case "Microsoft.Outlook.Mail.Read":
                case "Microsoft.Outlook.Explorer":
                    result = (Globals.ThisAddIn.OutlookVersion <= OutlookMajorVersion.Outlook2007) ?
                        GetResourceText("SuiteCRMAddIn.Menus.MailRead.xml") :
                        GetResourceText("SuiteCRMAddIn.Menus.MailRead2007.xml");
                    break;
                case "Microsoft.Outlook.Mail.Compose":
                    result = GetResourceText("SuiteCRMAddIn.Menus.MailCompose.xml");
                    break;
                default:
                    result = String.Empty;
                    break;
            }

            return result;
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            Globals.ThisAddIn.RibbonUI = ribbonUI;
        }

        #endregion

        public bool btnArchive_Enabled()
        {
            return Globals.ThisAddIn.HasCrmUserSession &&
//                Globals.ThisAddIn.Application.ActiveInspector().CurrentItem is Outlook.MailItem &&
                Globals.ThisAddIn.SelectedEmails.Select(x => x.UserProperties[SyncStateManager.CrmIdPropertyName] == null).Any();
        }

        public bool manualSyncButton_Enabled(IRibbonControl control)
        {
            return Globals.ThisAddIn.HasCrmUserSession &&
                   Globals.ThisAddIn.SelectedContacts.Count() == 1 &&
                   Settings.Default.SyncContacts == SyncDirection.Direction.Neither;
        }

        #region Click Events
        public void btnArchive_Action(IRibbonControl control)
        {
            DoOrLogError(() => Globals.ThisAddIn.ShowArchiveForm());
        }

        public void manualSyncButton_Action(IRibbonControl control)
        {
            DoOrLogError(() => Globals.ThisAddIn.ManualSyncContact());
        }

        public void btnSettings_Action(IRibbonControl control)
        {
            DoOrLogError(() =>
                Globals.ThisAddIn.ShowSettingsForm());
        }

        public void btnAddressBook_Action(IRibbonControl control)
        {
            DoOrLogError(() => Globals.ThisAddIn.ShowAddressBook());
        }

        /// <summary>
        /// Send, and also archive to CRM, the current message in the composer window.
        /// </summary>
        /// <param name="control">The ribbon which caused this action to be raised.</param>
        public void btnSendAndArchive_Action(IRibbonControl control)
        {
            Outlook.MailItem olItem = 
                (Globals.ThisAddIn.Application.ActiveInspector().CurrentItem as Outlook.MailItem);

            if (olItem != null)
            {
                if (Globals.ThisAddIn.HasCrmUserSession)
                {
                    try
                    {
                        try
                        {
                            List<Outlook.MailItem> items = new List<Outlook.MailItem>();
                            items.Add(olItem);

                            if ( new ArchiveDialog(items, EmailArchiveReason.SendAndArchive).ShowDialog() == DialogResult.OK)
                            {
                                olItem.Send();
                            }
                        }
                        catch (Exception failedToArchve)
                        {
                            Globals.ThisAddIn.ShowAndLogError(
                                failedToArchve,
                                $"Failed to archive message because {failedToArchve.Message}",
                                "Failed to archive");
                        }
                    }
                    catch (Exception failedToSend)
                    {
                        Globals.ThisAddIn.ShowAndLogError(
                            failedToSend, 
                            $"Failed to send message because {failedToSend.Message}", 
                            "Failed to send");
                    }
                }
                else
                {
                    ShowNoSessionWarning();
                }
            }
            else
            {
                Globals.ThisAddIn.Log.AddEntry(
                    "No message while attempting to send and archive?", 
                    LogEntryType.Warning);
            }
        }

        private static void ShowNoSessionWarning()
        {
            MessageBox.Show(
                "Please check your CRM login credentials in Settings and retry.",
                "No CRM Session",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private void ManualArchive()
        {
            if (Globals.ThisAddIn.SuiteCRMUserSession.NotLoggedIn)
            {
                Globals.ThisAddIn.ShowSettingsForm();
            }
            Globals.ThisAddIn.ShowArchiveForm();
        }
        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        /// <summary>
        /// Wrapper around invocation of an action, to provide consistent logging of 
        /// otherwise-uncaught exceptions.
        /// </summary>
        /// <param name="action">The actual action handler to invoke.</param>
        private void DoOrLogError(Action action)
        {
            Robustness.DoOrLogError(Globals.ThisAddIn.Log, action);
        }
    }
}
