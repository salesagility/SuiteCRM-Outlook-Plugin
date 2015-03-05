/**
 * Outlook integration for SuiteCRM.
 * @package Outlook integration for SuiteCRM
 * @copyright SalesAgility Ltd http://www.salesagility.com
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU AFFERO GENERAL PUBLIC LICENSE as published by
 * the Free Software Foundation; either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU AFFERO GENERAL PUBLIC LICENSE
 * along with this program; if not, see http://www.gnu.org/licenses
 * or write to the Free Software Foundation,Inc., 51 Franklin Street,
 * Fifth Floor, Boston, MA 02110-1301  USA
 *
 * @author SalesAgility <info@salesagility.com>
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using SuiteCRMClient;
using System.Runtime.InteropServices;
using SuiteCRMAddIn.Properties;
using System.Security.Cryptography;
using System.Globalization;
using SuiteCRMClient.RESTObjects;
using SuiteCRMClient;
using System.IO;
using System.Windows.Forms;

namespace SuiteCRMAddIn
{
    public partial class ThisAddIn
    {
        public SuiteCRMClient.clsUsersession SuiteCRMUserSession;
        public clsSettings settings;
        private Outlook.Explorer objExplorer;
        public Office.CommandBarPopup objSuiteCRMMenuBar2007;
        public Office.CommandBarButton btnArvive;
        public Office.CommandBarButton btnSettings;
        List<Outlook.Folder> lstOutlookFolders;
        public int CurrentVersion;
        public Office.IRibbonUI RibbonUI { get; set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            CurrentVersion = Convert.ToInt32(Globals.ThisAddIn.Application.Version.Split('.')[0]);
            this.objExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
            SuiteCRMClient.clsSuiteCRMHelper.InstallationPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SuiteCRMOutlookAddIn";
            this.settings = new clsSettings();
            if (this.settings.AutoArchive)
            {
                this.objExplorer.Application.NewMailEx += new Outlook.ApplicationEvents_11_NewMailExEventHandler(this.Application_NewMail);
                this.objExplorer.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(this.Application_ItemSend);
            }
            if (CurrentVersion < 14)
            {
                this.Application.ItemContextMenuDisplay += new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(this.Application_ItemContextMenuDisplay);
                var menuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;
                objSuiteCRMMenuBar2007 = (Office.CommandBarPopup)menuBar.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, true);
                if (objSuiteCRMMenuBar2007 != null)
                {
                    objSuiteCRMMenuBar2007.Caption = "SuiteCRM";
                    this.btnArvive = (Office.CommandBarButton)this.objSuiteCRMMenuBar2007.Controls.Add(Office.MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                    this.btnArvive.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    this.btnArvive.Caption = "Archive";
                    this.btnArvive.Picture = RibbonImageHelper.Convert(Resources.SuiteCRM1);
                    this.btnArvive.Click += new Office._CommandBarButtonEvents_ClickEventHandler(this.cbtnArchive_Click);
                    this.btnArvive.Visible = true;
                    this.btnArvive.BeginGroup = true;
                    this.btnArvive.TooltipText = "Archive selected emails to SuiteCRM";
                    this.btnArvive.Enabled = true;
                    this.btnSettings = (Office.CommandBarButton)this.objSuiteCRMMenuBar2007.Controls.Add(Office.MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                    this.btnSettings.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    this.btnSettings.Caption = "Settings";
                    this.btnSettings.Click += new Office._CommandBarButtonEvents_ClickEventHandler(this.cbtnSettings_Click);
                    this.btnSettings.Visible = true;
                    this.btnSettings.BeginGroup = true;
                    this.btnSettings.TooltipText = "SuiteCRM Settings";
                    this.btnSettings.Enabled = true;
                    this.btnSettings.Picture = RibbonImageHelper.Convert(Resources.Settings);            

                    objSuiteCRMMenuBar2007.Visible = true;
                }
            }
            else
            {
                //For Outlook version 2010 and greater
                //var app = this.Application;
                //app.FolderContextMenuDisplay += new Outlook.ApplicationEvents_11_FolderContextMenuDisplayEventHander(this.app_FolderContextMenuDisplay);
            }
            SuiteCRMAuthenticate();
        }

        //void app_FolderContextMenuDisplay(Office.CommandBar CommandBar, Outlook.MAPIFolder Folder)
        //{
        //    RibbonUI.InvalidateControlMso("FolderPropertiesContext");
        //} 
        private void cbtnArchive_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ManualArchive();
        }

        private void cbtnSettings_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            frmSettings objacbbSettings = new frmSettings();
            objacbbSettings.ShowDialog();
        }

        private void ManualArchive()
        {
            if (Globals.ThisAddIn.SuiteCRMUserSession.id == "")
            {
                frmSettings objacbbSettings = new frmSettings();
                objacbbSettings.ShowDialog();
            }
            if (Globals.ThisAddIn.SuiteCRMUserSession.id != "")
            {
                frmArchive objForm = new frmArchive();
                objForm.ShowDialog();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                if (SuiteCRMUserSession != null)
                    SuiteCRMUserSession.LogOut();
                if (this.CommandBarExists("SuiteCRM"))
                {
                    this.objSuiteCRMMenuBar2007.Delete();
                }
                this.UnregisterEvents();
            }
            catch (Exception ex)
            {
                ex.Data.Clear();
            }
        }

        private void UnregisterEvents()
        {
            try
            {
                this.Application.ItemContextMenuDisplay -= new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(this.Application_ItemContextMenuDisplay);
            }
            catch (System.Exception ex)
            {

            }

            try
            {
                this.btnArvive.Click -= new Office._CommandBarButtonEvents_ClickEventHandler(this.cbtnArchive_Click);
            }
            catch (System.Exception ex1)
            {

            }

            try
            {
                this.objExplorer.Application.NewMailEx -= new Outlook.ApplicationEvents_11_NewMailExEventHandler(this.Application_NewMail);
            }
            catch (System.Exception ex2)
            {

            }

            try
            {
                this.objExplorer.Application.ItemSend -= new Outlook.ApplicationEvents_11_ItemSendEventHandler(this.Application_ItemSend);
            }
            catch (System.Exception ex3)
            {

            }


        }

        private bool CommandBarExists(string name)
        {
            try
            {
                string text1 = Globals.ThisAddIn.Application.ActiveExplorer().CommandBars[name].Name;
                return true;
            }
            catch (System.Exception)
            {
                return false;
            }
        }

        #region VSTO generated code

        private void Application_ItemContextMenuDisplay(Office.CommandBar CommandBar, Outlook.Selection Selection)
        {
            Outlook.Selection selection = Selection;
            Outlook.MailItem item1 = (Outlook.MailItem)selection[1];
            Office.CommandBarButton objMainMenu = (Office.CommandBarButton)CommandBar.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, this.missing, this.missing, this.missing, this.missing);
            objMainMenu.Caption = "SuiteCRM Archive";
            objMainMenu.Visible = true;
            objMainMenu.Picture = RibbonImageHelper.Convert(Resources.SuiteCRM1);
            objMainMenu.Click += new Office._CommandBarButtonEvents_ClickEventHandler(this.contextMenuArchiveButton_Click);
        }

        private void contextMenuArchiveButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (Globals.ThisAddIn.SuiteCRMUserSession.id == "")
            {
                frmSettings objacbbSettings = new frmSettings();
                objacbbSettings.ShowDialog();
            }
            frmArchive objForm = new frmArchive();
            objForm.ShowDialog();
        }

        private void Application_ItemSend(object item, ref bool target)
        {
            try
            {
                if (item is Outlook.MailItem)
                {
                    Outlook.MailItem objMail = (Outlook.MailItem)item;
                    if (objMail.UserProperties["SuiteCRM"] == null)
                    {
                        ArchiveEmail(objMail, 3, this.settings.ExcludedEmails);
                        objMail.UserProperties.Add("SuiteCRM", Outlook.OlUserPropertyType.olText, true, Outlook.OlUserPropertyType.olText);
                        objMail.UserProperties["SuiteCRM"].Value = "True";
                        objMail.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "OutlookEvents_ItemSend General Exception:" + "\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "Data:" + ex.Data.ToString() + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------" + "\n";
                clsSuiteCRMHelper.WriteLog(strLog);
            }
        }

        private void Application_NewMail(string EntryID)
        {
            try
            {
                var objItem = Globals.ThisAddIn.Application.Session.GetItemFromID(EntryID);
                if (objItem is Outlook.MailItem)
                {
                    Outlook.MailItem objMail = (Outlook.MailItem)objItem;
                    if (objMail.UserProperties["SuiteCRM"] == null)
                    {
                        ArchiveEmail(objMail, 2, this.settings.ExcludedEmails);
                        objMail.UserProperties.Add("SuiteCRM", Outlook.OlUserPropertyType.olText, true, Outlook.OlUserPropertyType.olText);
                        objMail.UserProperties["SuiteCRM"].Value = "True";
                        objMail.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "Application_NewMail General Exception:" + "\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "Data:" + ex.Data.ToString() + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------" + "\n";
                clsSuiteCRMHelper.WriteLog(strLog);
            }
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new SuiteCRMRibbon();
        }

        public void SuiteCRMAuthenticate()
        {
            if (Globals.ThisAddIn.SuiteCRMUserSession == null)
            {
                Authenticate();
            }
            else
            {
                if (Globals.ThisAddIn.SuiteCRMUserSession.id == "")
                    Authenticate();
            }

        }

        public void Authenticate()
        {
            try
            {
                string strUsername = Globals.ThisAddIn.settings.username;
                string strPassword = Globals.ThisAddIn.settings.password;                

                Globals.ThisAddIn.SuiteCRMUserSession = new SuiteCRMClient.clsUsersession("", "", "","");
                string strURL = Globals.ThisAddIn.settings.host;
                if (strURL != "")
                {
                    Globals.ThisAddIn.SuiteCRMUserSession = new SuiteCRMClient.clsUsersession(strURL, strUsername, strPassword, Globals.ThisAddIn.settings.LDAPKey);
                    Globals.ThisAddIn.SuiteCRMUserSession.AwaitingAuthentication = true;
                    try
                    {
                        if (settings.IsLDAPAuthentication)
                        {
                            Globals.ThisAddIn.SuiteCRMUserSession.AuthenticateLDAP();
                        }
                        else
                            Globals.ThisAddIn.SuiteCRMUserSession.Login();
                        if (Globals.ThisAddIn.SuiteCRMUserSession.id != "")
                            return;
                    }
                    catch (Exception ex)
                    {
                        ex.Data.Clear();
                    }
                }
                Globals.ThisAddIn.SuiteCRMUserSession.AwaitingAuthentication = false;
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "Authenticate method General Exception:\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------\n";
                clsSuiteCRMHelper.WriteLog(strLog);
                ex.Data.Clear();
            }
        }

        private void GetMailFolders(Outlook.Folders objInpFolders)
        {
            try
            {
                foreach (Outlook.Folder objFolder in objInpFolders)
                {
                    if (objFolder.Folders.Count > 0)
                    {
                        lstOutlookFolders.Add(objFolder);
                        GetMailFolders(objFolder.Folders);
                    }
                    else
                        lstOutlookFolders.Add(objFolder);
                }
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "GetMailFolders method General Exception:" + "\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "Data:" + ex.Data.ToString() + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------" + "\n";
                clsSuiteCRMHelper.WriteLog(strLog);
                ex.Data.Clear();
            }
        }

        private void ArchiveEmail(Outlook.MailItem objMail, int intArchiveType, string strExcludedEmails = "")
        {
            try
            {
                SuiteCRMClient.clsEmailArchive objEmail = new SuiteCRMClient.clsEmailArchive();
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
                objEmail.ArchiveType = intArchiveType;
                foreach (Outlook.Attachment objMailAttachments in objMail.Attachments)
                {
                    objEmail.Attachments.Add(new SuiteCRMClient.clsEmailAttachments { DisplayName = objMailAttachments.DisplayName, FileContentInBase64String = Base64Encode(objMailAttachments, objMail) });
                }

                System.Threading.Thread objThread = new System.Threading.Thread(() => ArchiveEmailThread(objEmail, intArchiveType, strExcludedEmails));
                objThread.Start();
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "ArchiveEmail method General Exception:" + "\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "Data:" + ex.Data.ToString() + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------" + "\n";
                clsSuiteCRMHelper.WriteLog(strLog);
                ex.Data.Clear();
            }
        }

        private void ArchiveEmailThread(SuiteCRMClient.clsEmailArchive objEmail, int intArchiveType, string strExcludedEmails = "")
        {
            try
            {
                if (SuiteCRMUserSession != null)
                {
                    while (SuiteCRMUserSession.AwaitingAuthentication == true)
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    if (SuiteCRMUserSession.id != "")
                    {
                        objEmail.SuiteCRMUserSession = SuiteCRMUserSession;
                        objEmail.Save(strExcludedEmails);
                    }
                }
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "ArchiveEmailThread method General Exception:" + "\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "Data:" + ex.Data.ToString() + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------" + "\n";
                clsSuiteCRMHelper.WriteLog(strLog);
                ex.Data.Clear();
            }

        }

        public byte[] Base64Encode(Outlook.Attachment objMailAttachment, Outlook.MailItem objMail)
        {
            byte[] strRet = null;
            if (objMailAttachment != null)
            {
                if (System.IO.Directory.Exists(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath") == false)
                {
                    string strPath = Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath";
                    System.IO.Directory.CreateDirectory(strPath);
                }
                try
                {
                    objMailAttachment.SaveAsFile(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath\\" + objMailAttachment.FileName);
                    strRet = System.IO.File.ReadAllBytes(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath\\" + objMailAttachment.FileName);
                }
                catch (COMException ex)
                {
                    try
                    {
                        string strLog;
                        strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                        strLog += "AddInModule.Base64Encode method COM Exception:" + "\n";
                        strLog += "Message:" + ex.Message + "\n";
                        strLog += "Source:" + ex.Source + "\n";
                        strLog += "StackTrace:" + ex.StackTrace + "\n";
                        strLog += "Data:" + ex.Data.ToString() + "\n";
                        strLog += "HResult:" + ex.HResult.ToString() + "\n";
                        strLog += "Inputs:" + "\n";
                        strLog += "Data:" + objMailAttachment.DisplayName + "\n";
                        strLog += "-------------------------------------------------------------------------" + "\n";
                        clsSuiteCRMHelper.WriteLog(strLog);
                        ex.Data.Clear();
                        string strName = Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath\\" + DateTime.Now.ToString("MMddyyyyHHmmssfff") + ".html";
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
                        string strLog;
                        strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                        strLog += "AddInModule.Base64Encode method General Exception:" + "\n";
                        strLog += "Message:" + ex.Message + "\n";
                        strLog += "Source:" + ex.Source + "\n";
                        strLog += "StackTrace:" + ex.StackTrace + "\n";
                        strLog += "Data:" + ex.Data.ToString() + "\n";
                        strLog += "HResult:" + ex.HResult.ToString() + "\n";
                        strLog += "Inputs:" + "\n";
                        strLog += "Data:" + objMailAttachment.DisplayName + "\n";
                        strLog += "-------------------------------------------------------------------------" + "\n";
                        clsSuiteCRMHelper.WriteLog(strLog);
                        ex1.Data.Clear();
                    }
                }
                finally
                {
                    if (System.IO.Directory.Exists(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath") == true)
                    {
                        System.IO.Directory.Delete(Environment.SpecialFolder.MyDocuments.ToString(), true);
                    }
                }
            }

            return strRet;
        }

        private void ArchiveFolderItems(Outlook.Folder objFolder, DateTime? dtAutoArchiveFrom = null)
        {
            try
            {
                Outlook.Items UnReads;
                if (dtAutoArchiveFrom == null)
                    UnReads = objFolder.Items.Restrict("[Unread]=true");
                else
                    UnReads = objFolder.Items.Restrict("[ReceivedTime] >= '" + ((DateTime)dtAutoArchiveFrom).AddDays(-1).ToString("yyyy-MM-dd HH:mm") + "'");

                for (int intItr = 1; intItr <= UnReads.Count; intItr++)
                {
                    if (UnReads[intItr] is Outlook.MailItem)
                    {
                        Outlook.MailItem objMail = (Outlook.MailItem)UnReads[intItr];

                        if (objMail.UserProperties["SuiteCRM"] == null)
                        {
                            ArchiveEmail(objMail, 2, this.settings.ExcludedEmails);
                            objMail.UserProperties.Add("SuiteCRM", Outlook.OlUserPropertyType.olText, true, Outlook.OlUserPropertyType.olText);
                            objMail.UserProperties["SuiteCRM"].Value = "True";
                            objMail.Save();
                        }
                        else
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "ArchiveFolderItems method General Exception:" + "\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "Data:" + ex.Data.ToString() + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------" + "\n";
                clsSuiteCRMHelper.WriteLog(strLog);
                ex.Data.Clear();
            }
        }

        public void ProcessMails(DateTime? dtAutoArchiveFrom = null)
        {
            if (settings.AutoArchive == false)
                return;
            System.Threading.Thread.Sleep(5000);
            while (1 == 1)
            {
                try
                {
                    lstOutlookFolders = new List<Outlook.Folder>();
                    GetMailFolders(Globals.ThisAddIn.Application.Session.Folders);
                    if (lstOutlookFolders != null)
                    {
                        foreach (Outlook.Folder objFolder in lstOutlookFolders)
                        {
                            if (settings.AutoArchiveFolders == null)
                                ArchiveFolderItems(objFolder, dtAutoArchiveFrom);
                            else if (settings.AutoArchiveFolders.Count == 0)
                                ArchiveFolderItems(objFolder, dtAutoArchiveFrom);
                            else
                            {
                                if (settings.AutoArchiveFolders.Contains(objFolder.EntryID))
                                {
                                    ArchiveFolderItems(objFolder, dtAutoArchiveFrom);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    string strLog;
                    strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                    strLog += "ProcessMails method General Exception:" + "\n";
                    strLog += "Message:" + ex.Message + "\n";
                    strLog += "Source:" + ex.Source + "\n";
                    strLog += "StackTrace:" + ex.StackTrace + "\n";
                    strLog += "Data:" + ex.Data.ToString() + "\n";
                    strLog += "HResult:" + ex.HResult.ToString() + "\n";
                    strLog += "-------------------------------------------------------------------------" + "\n";
                    clsSuiteCRMHelper.WriteLog(strLog);
                    ex.Data.Clear();
                }
                if (dtAutoArchiveFrom != null)
                    break;

                System.Threading.Thread.Sleep(5000);
            }
        }        
    }
}
