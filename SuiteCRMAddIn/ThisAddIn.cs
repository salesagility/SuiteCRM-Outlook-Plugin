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
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using SuiteCRMClient;
using System.Runtime.InteropServices;
using SuiteCRMAddIn.Properties;
using System.Globalization;
using SuiteCRMClient.RESTObjects;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using SuiteCRMClient.Logging;

namespace SuiteCRMAddIn
{
    using System.Reflection;
    using BusinessLogic;
    using SuiteCRMClient.Email;

    public partial class ThisAddIn
    {
        public static readonly string AddInTitle, AddInVersion;

        public SuiteCRMClient.clsUsersession SuiteCRMUserSession;
        public clsSettings settings;
        private Outlook.Explorer objExplorer;
        public Office.CommandBarPopup objSuiteCRMMenuBar2007;
        public Office.CommandBarButton btnArvive;
        public Office.CommandBarButton btnSettings;
        public int OutlookVersion;

        private SyncContext _syncContext;
        private ContactSyncing _contactSyncing;
        private TaskSyncing _taskSyncing;
        private CalendarSyncing _calendarSyncing;

        public Office.IRibbonUI RibbonUI { get; set; }

        public ILogger Log;

        static ThisAddIn()
        {
            GetTitleAndVersion(out AddInTitle, out AddInVersion);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                var outlookApp = this.Application;
                OutlookVersion = Convert.ToInt32(outlookApp.Version.Split('.')[0]);

                this.settings = new clsSettings();
                _syncContext = new SyncContext(outlookApp, settings);
                _contactSyncing = new ContactSyncing(_syncContext);
                _taskSyncing = new TaskSyncing(_syncContext);
                _calendarSyncing = new CalendarSyncing(_syncContext);

                var outlookExplorer = outlookApp.ActiveExplorer();
                this.objExplorer = outlookExplorer;
                StartLogging(settings);
                outlookExplorer.FolderSwitch -= objExplorer_FolderSwitch;
                outlookExplorer.FolderSwitch += objExplorer_FolderSwitch;
                
                // TODO: install/remove these event handlers when settings.AutoArchive changes:
                outlookApp.NewMailEx += new Outlook.ApplicationEvents_11_NewMailExEventHandler(this.Application_NewMail);
                outlookApp.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(this.Application_ItemSend);

                if (OutlookVersion < 14)
                {
                    outlookApp.ItemContextMenuDisplay += new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(this.Application_ItemContextMenuDisplay);
                    var menuBar = outlookExplorer.CommandBars.ActiveMenuBar;
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
                Sync();
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.ThisAddIn_Startup", ex);
            }
        }

        public static string LogDirPath =>
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
            "\\SuiteCRMOutlookAddIn\\Logs\\";

        private void StartLogging(clsSettings settings)
        {
            Log = Log4NetLogger.FromFilePath("add-in", LogDirPath + "suitecrmoutlook.log", () => GetLogHeader(settings));
            clsSuiteCRMHelper.SetLog(Log);
            SuiteCRMClient.CrmRestServer.SetLog(Log);
        }

        private void LogKeySettings(clsSettings settings)
        {
            foreach (var s in GetKeySettings(settings))
            {
                Log.Info(s);
            }
        }

        private IEnumerable<string> GetLogHeader(clsSettings settings)
        {
            yield return $"{AddInTitle} v{AddInVersion}";
            foreach (var s in GetKeySettings(settings)) yield return s;
        }

        private IEnumerable<string> GetKeySettings(clsSettings settings)
        {
            yield return "Auto-archiving: " + (settings.AutoArchive ? "ON" : "off");
        }

        void objExplorer_FolderSwitch()
        {
            try
            {
                _syncContext.SetCurrentFolder(this.objExplorer.CurrentFolder);
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.objExplorer_FolderSwitch", ex);
            }
        }

        public async void Sync()
        //public void Sync()
        {
            try
            {
                while (true)
                {
                    if (SuiteCRMUserSession != null && SuiteCRMUserSession.id != "")
                    {
                        if (settings.SyncCalendar)
                        {
                            //StartCalendarSync();
                            // for test !!!
                            Thread oThread = new Thread(() =>_calendarSyncing.StartCalendarSync());
                            oThread.Start();
                            //StartTaskSync();
                            Thread oThread1 = new Thread(() => _taskSyncing.StartTaskSync());
                            oThread1.Start();
                        }
                        if (settings.SyncContacts)
                        {
                            //StartContactSync();
                            Thread oThread = new Thread(() => _contactSyncing.StartContactSync());
                            oThread.Start();
                        }
                    }
                    await Task.Delay(300000); //5 mins delay
                }
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.Sync", ex);
            }
        }
        private void cbtnArchive_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ManualArchive();
        }

        private void cbtnSettings_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ShowSettingsForm();
        }

        public void ShowSettingsForm()
        {
            var settingsForm = new frmSettings();
            settingsForm.SettingsChanged += (sender, args) => this.LogKeySettings(settings);
            settingsForm.ShowDialog();
        }

        public void ShowArchiveForm()
        {
            frmArchive objForm = new frmArchive();
            objForm.ShowDialog();
        }

        private void ManualArchive()
        {
            if (Globals.ThisAddIn.SuiteCRMUserSession.id == "")
            {
                ShowSettingsForm();
            }
            if (Globals.ThisAddIn.SuiteCRMUserSession.id != "")
            {
                ShowArchiveForm();
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
                Log.Error("ThisAddIn.ThisAddIn_Shutdown", ex);
            }
        }

        private void UnregisterEvents()
        {
            try
            {
                this.Application.ItemContextMenuDisplay -= new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(this.Application_ItemContextMenuDisplay);
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.UnregisterEvents", ex);
            }
            try
            {
                this.btnArvive.Click -= new Office._CommandBarButtonEvents_ClickEventHandler(this.cbtnArchive_Click);
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.UnregisterEvents", ex);
            }

            try
            {
                this.objExplorer.Application.NewMailEx -= new Outlook.ApplicationEvents_11_NewMailExEventHandler(this.Application_NewMail);
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.UnregisterEvents", ex);
            }

            try
            {
                this.objExplorer.Application.ItemSend -= new Outlook.ApplicationEvents_11_ItemSendEventHandler(this.Application_ItemSend);
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.UnregisterEvents", ex);
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
            try
            {
                Outlook.Selection selection = Selection;
                Outlook.MailItem item1 = (Outlook.MailItem)selection[1];
                Office.CommandBarButton objMainMenu = (Office.CommandBarButton)CommandBar.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, this.missing, this.missing, this.missing, this.missing);
                objMainMenu.Caption = "SuiteCRM Archive";
                objMainMenu.Visible = true;
                objMainMenu.Picture = RibbonImageHelper.Convert(Resources.SuiteCRM1);
                objMainMenu.Click += new Office._CommandBarButtonEvents_ClickEventHandler(this.contextMenuArchiveButton_Click);
            }
            catch (Exception)
            {

            }
        }

        private void contextMenuArchiveButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (Globals.ThisAddIn.SuiteCRMUserSession.id == "")
            {
                ShowSettingsForm();
            }
            ShowArchiveForm();
        }

        private void Application_ItemSend(object item, ref bool target)
        {
            try
            {
                if (!settings.AutoArchive) return;
                ProcessNewMailItem(EmailArchiveType.Sent, item as Outlook.MailItem);
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.Application_ItemSend", ex);
            }
        }

        private void Application_NewMail(string EntryID)
        {
            try
            {
                if (!settings.AutoArchive) return;
                ProcessNewMailItem(
                    EmailArchiveType.Inbound,
                    Application.Session.GetItemFromID(EntryID) as Outlook.MailItem);
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.Application_NewMail", ex);
            }
        }

        private void ProcessNewMailItem(EmailArchiveType archiveType, Outlook.MailItem mailItem)
        {
            if (mailItem == null)
            {
                Log.Info("New 'mail item' was null");
                return;
            }
            Log.Info("Processing new mail item: " + mailItem.Subject);
            new EmailArchiving().ProcessEligibleNewMailItem(mailItem, archiveType);
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

                Globals.ThisAddIn.SuiteCRMUserSession = new SuiteCRMClient.clsUsersession("", "", "","", Log);
                string strURL = Globals.ThisAddIn.settings.host;
                if (strURL != "")
                {
                    Globals.ThisAddIn.SuiteCRMUserSession = new SuiteCRMClient.clsUsersession(strURL, strUsername, strPassword, Globals.ThisAddIn.settings.LDAPKey, Log);
                    Globals.ThisAddIn.SuiteCRMUserSession.AwaitingAuthentication = true;
                    try
                    {
                        if (settings.IsLDAPAuthentication)
                        {
                            Globals.ThisAddIn.SuiteCRMUserSession.AuthenticateLDAP();
                        }
                        else
                        {
                            Globals.ThisAddIn.SuiteCRMUserSession.Login();
                        }

                        if (Globals.ThisAddIn.SuiteCRMUserSession.id != "")
                            return;
                    }
                    catch (Exception)
                    {
                        // Swallow exception(!)
                    }
                }
                Globals.ThisAddIn.SuiteCRMUserSession.AwaitingAuthentication = false;
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.Authenticate", ex);
            }
        }

        public void ProcessMails(DateTime? dtAutoArchiveFrom = null)
        {
            new EmailArchiving().ProcessMails(dtAutoArchiveFrom);
        }

        public int SelectedEmailCount => Application.ActiveExplorer()?.Selection.Count ?? 0;

        public IEnumerable<Outlook.MailItem> SelectedEmails
        {
            get
            {
                var selection = Application.ActiveExplorer()?.Selection;
                if (selection == null) yield break;
                foreach (object e in selection)
                {
                    var mail = e as Outlook.MailItem;
                    if (mail != null) yield return mail;
                    Marshal.ReleaseComObject(e);
                }
            }
        }

        private static void GetTitleAndVersion(out string title, out string versionString)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var name = assembly.GetName();
            var version = name.Version;

            title = (assembly.GetCustomAttributes(typeof(AssemblyTitleAttribute)).SingleOrDefault() as
                        AssemblyTitleAttribute)?.Title
                    ?? name.Name;

            // 'Build' is what we'd call the 'revision number' and
            // 'Revision' is what we'd call the 'build number'.
            versionString = $"{version.Major}.{version.Minor}.{version.Build}.{version.Revision}";
        }
    }
}
