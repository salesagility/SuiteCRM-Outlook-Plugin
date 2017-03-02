﻿/**
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
    using BusinessLogic;
    using SuiteCRMAddIn.Properties;
    using SuiteCRMClient;
    using SuiteCRMClient.Email;
    using SuiteCRMClient.Logging;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using Office = Microsoft.Office.Core;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public partial class ThisAddIn
    {
        public static readonly string AddInTitle, AddInVersion;

        public SuiteCRMClient.UserSession SuiteCRMUserSession;
        private clsSettings settings;
        private Outlook.Explorer objExplorer;
        public Office.CommandBarPopup objSuiteCRMMenuBar2007;
        public Office.CommandBarButton btnArvive;
        public Office.CommandBarButton btnSettings;
        public OutlookMajorVersion OutlookVersion;

        private SyncContext synchronisationContext;
        private ContactSyncing contactSynchroniser;
        private TaskSyncing taskSynchroniser;
        private AppointmentSyncing appointmentSynchroniser;

        public Office.IRibbonUI RibbonUI { get; set; }
        private EmailArchiving emailArchiver;

        private ILogger log;

        /// <summary>
        /// Property to allow other classes to get the logger object, but not
        /// replace it with a new logger object.
        /// </summary>
        public ILogger Log
        {
            get { return this.log; }
        }

        /// <summary>
        /// Property to allow other classes to get the settings object, but not
        /// replace it with a new settings object.
        /// </summary>
        public clsSettings Settings
        {
            get
            {
                return this.settings;
            }
        }

        /// <summary>
        /// I'm guessing this method is called once when the add-in is added in.
        /// </summary>
        static ThisAddIn()
        {
            GetTitleAndVersion(out AddInTitle, out AddInVersion);
        }

        /// <summary>
        /// Make a call out to store.suitecrm.com with the add-in's public key (global to all 
        /// instances of the plugin, and stashed probably on the assembly), and this instance's
        /// key (specific to this instance, and entered by the user through the settings form).
        /// </summary>
        /// <returns>true if licence key was verified, or if licence server could not be reached.</returns>
        private bool VerifyLicenceKey()
        {
            bool result = false;
            try
            {
                result = new LicenceValidationHelper(this.Log, Properties.Settings.Default.PublicKey, this.settings.LicenceKey).Validate();
            } catch (System.Configuration.SettingsPropertyNotFoundException ex)
            {
                this.log.Error("Licence key was not yet set up", ex);
            }
            return result;
        }

        public bool HasCrmUserSession
        {
            get { return SuiteCRMUserSession?.IsLoggedIn ?? false; }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Prepare();

                Run();
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.ThisAddIn_Startup", ex);
            }
        }

        /// <summary>
        /// Prepare me for running.
        /// </summary>
        private void Prepare()
        {
            var outlookApp = this.Application;
            OutlookVersion = (OutlookMajorVersion)Convert.ToInt32(outlookApp.Version.Split('.')[0]);

            this.settings = new clsSettings();
            StartLogging(settings);

            synchronisationContext = new SyncContext(outlookApp, settings);
            contactSynchroniser = new ContactSyncing("CS", synchronisationContext);
            taskSynchroniser = new TaskSyncing("TS", synchronisationContext);
            appointmentSynchroniser = new AppointmentSyncing("AS", synchronisationContext);
            emailArchiver = new EmailArchiving("EM", synchronisationContext.Log);

            var outlookExplorer = outlookApp.ActiveExplorer();
            this.objExplorer = outlookExplorer;
            outlookExplorer.FolderSwitch -= objExplorer_FolderSwitch;
            outlookExplorer.FolderSwitch += objExplorer_FolderSwitch;

            // TODO: install/remove these event handlers when settings.AutoArchive changes:
            outlookApp.NewMailEx += new Outlook.ApplicationEvents_11_NewMailExEventHandler(this.Application_NewMail);
            outlookApp.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(this.Application_ItemSend);

            if (OutlookVersion < OutlookMajorVersion.Outlook2010)
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
        }

        /// <summary>
        /// Check the licence; if valid, do normal processing; otherwise, give the user the options of 
        /// reconfiguring or disabling the add-in.
        /// </summary>
        private void Run()
        {
            bool success = false, disable = false;
            for (success = false; !(success || disable);)
            {
                success = this.VerifyLicenceKey();

                if (!success)
                {
                    switch (new ReconfigureOrDisableDialog().ShowDialog())
                    {
                        case DialogResult.OK:
                            /* if licence key does not validate, show the settings form to allow the user to enter
                             * a (new) key, and retry. */
                            Log.Info("User chose to reconfigure add-in");
                            this.ShowSettingsForm();
                            break;
                        case DialogResult.Cancel:
                            Log.Info("User chose to disable add-in");
                            disable = true;
                            break;
                        default:
                            log.Warn("Unexpected response from ReconfigureOrDisableDialog");
                            disable = true;
                            break;
                    }
                }
            }

            if (success)
            {
                log.Info("Licence verified, starting normal operation.");
                SuiteCRMAuthenticate();
                StartSynchronisationProcesses();
            }
            else /* presumably disable is true */
            {
                Log.Warn("Disabling add-in");
                Application.COMAddIns.Item("SuiteCRMAddIn").Connect = false;
            }
        }

        public static string LogDirPath =>
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
            "\\SuiteCRMOutlookAddIn\\Logs\\";

        private void StartLogging(clsSettings settings)
        {
            log = Log4NetLogger.FromFilePath("add-in", LogDirPath + "suitecrmoutlook.log", () => GetLogHeader(settings));
            clsSuiteCRMHelper.SetLog(log);
        }

        private void LogKeySettings(clsSettings settings)
        {
            foreach (var s in GetKeySettings(settings))
            {
                log.Info(s);
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
                synchronisationContext.SetCurrentFolder(this.objExplorer.CurrentFolder);
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.objExplorer_FolderSwitch", ex);
            }
        }

        public void StartSynchronisationProcesses()
        {
            DoOrLogError(() => this.appointmentSynchroniser.Start(), "Starting appointments synchroniser");
            DoOrLogError(() => this.contactSynchroniser.Start(), "Starting contacts synchroniser");
            DoOrLogError(() => this.taskSynchroniser.Start(), "Starting tasks synchroniser");
            DoOrLogError(() => this.emailArchiver.Start(), "Starting email archiver");
        }

        private void cbtnArchive_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ManualArchive();
        }

        private void cbtnSettings_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            DoOrLogError(() =>
                ShowSettingsForm());
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
            if (!HasCrmUserSession)
            {
                ShowSettingsForm();
            }
            if (HasCrmUserSession)
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
                log.Error("ThisAddIn.ThisAddIn_Shutdown", ex);
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
                log.Error("ThisAddIn.UnregisterEvents", ex);
            }
            try
            {
                this.btnArvive.Click -= new Office._CommandBarButtonEvents_ClickEventHandler(this.cbtnArchive_Click);
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.UnregisterEvents", ex);
            }

            try
            {
                this.objExplorer.Application.NewMailEx -= new Outlook.ApplicationEvents_11_NewMailExEventHandler(this.Application_NewMail);
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.UnregisterEvents", ex);
            }

            try
            {
                this.objExplorer.Application.ItemSend -= new Outlook.ApplicationEvents_11_ItemSendEventHandler(this.Application_ItemSend);
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.UnregisterEvents", ex);
            }

            try
            {
                appointmentSynchroniser.Dispose();
            }
            catch (Exception ex)
            {
                log.Error("AppointmentSyncing.Dispose", ex);
            }
            try
            {
                contactSynchroniser.Dispose();
            }
            catch (Exception ex)
            {
                log.Error("ContactSyncing.Dispose", ex);
            }
            try
            {
                taskSynchroniser.Dispose();
            }
            catch (Exception ex)
            {
                log.Error("TaskSyncing.Dispose", ex);
            }
        }

        private bool CommandBarExists(string name)
        {
            try
            {
                string text1 = Application.ActiveExplorer().CommandBars[name].Name;
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
            DoOrLogError(() =>
            {
                if (!HasCrmUserSession)
                {
                    ShowSettingsForm();
                }
                ShowArchiveForm();
            });
        }

        private void Application_ItemSend(object item, ref bool target)
        {
            log.Debug("Outlook ItemSend: email sent event");
            try
            {
                if (!settings.AutoArchive) return;
                ProcessNewMailItem(EmailArchiveType.Sent, item as Outlook.MailItem);
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.Application_ItemSend", ex);
            }
        }

        private void Application_NewMail(string EntryID)
        {
            log.Debug("Outlook NewMail: email received event");
            try
            {
                if (!settings.AutoArchive) return;
                ProcessNewMailItem(
                    EmailArchiveType.Inbound,
                    Application.Session.GetItemFromID(EntryID) as Outlook.MailItem);
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.Application_NewMail", ex);
            }
        }

        private void ProcessNewMailItem(EmailArchiveType archiveType, Outlook.MailItem mailItem)
        {
            if (mailItem == null)
            {
                log.Info("New 'mail item' was null");
                return;
            }
            new EmailArchiving($"EN-{mailItem.SenderEmailAddress}", Globals.ThisAddIn.Log).ProcessEligibleNewMailItem(mailItem, archiveType);
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
            if (!HasCrmUserSession)
            {
                Authenticate();
            }
        }

        public void Authenticate()
        {
            try
            {
                if (settings.host != String.Empty)
                {
                    SuiteCRMUserSession = 
                        new SuiteCRMClient.UserSession(
                            settings.host,
                            settings.username,
                            settings.password, 
                            settings.LDAPKey, 
                            log, 
                            Settings.RestTimeout);
                    SuiteCRMUserSession.AwaitingAuthentication = true;
                    try
                    {
                        if (settings.IsLDAPAuthentication)
                        {
                            SuiteCRMUserSession.AuthenticateLDAP();
                        }
                        else
                        {
                            SuiteCRMUserSession.Login();
                        }

                        if (SuiteCRMUserSession.IsLoggedIn)
                            return;
                    }
                    catch (Exception)
                    {
                        // Swallow exception(!)
                    }
                }
                else
                {
                    // We don't have a URL to connect to, dummy the connection.
                    SuiteCRMUserSession =
                        new SuiteCRMClient.UserSession(
                            String.Empty, 
                            String.Empty, 
                            String.Empty, 
                            String.Empty, 
                            log, 
                            Settings.RestTimeout);
                }

                SuiteCRMUserSession.AwaitingAuthentication = false;
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.Authenticate", ex);
            }
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

        private void DoOrLogError(Action action)
        {
            Robustness.DoOrLogError(log, action);
        }

        private void DoOrLogError(Action action, string message)
        {
            Robustness.DoOrLogError(log, action, message);
        }
    }
}
