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
    using BusinessLogic;
    using Daemon;
    using Dialogs;
    using Microsoft.Office.Core;
    using SuiteCRMAddIn.Properties;
    using SuiteCRMClient;
    using SuiteCRMClient.Email;
    using SuiteCRMClient.Logging;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Threading;
    using System.Windows.Forms;
    using Office = Microsoft.Office.Core;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public partial class ThisAddIn
    {
        private const string ProgId = "SuiteCRMAddIn";
        public static readonly string AddInTitle, AddInVersion;

        public SuiteCRMClient.UserSession SuiteCRMUserSession;
        private Outlook.Explorer objExplorer;
        public Office.CommandBarPopup objSuiteCRMMenuBar2007;
        public Office.CommandBarButton btnArchive;
        public Office.CommandBarButton btnSettings;
        public OutlookMajorVersion OutlookVersion;

        private SyncContext synchronisationContext;
        private ContactSyncing contactSynchroniser;
        private TaskSyncing taskSynchroniser;
        private AppointmentSyncing appointmentSynchroniser;

        public Office.IRibbonUI RibbonUI { get; set; }
        public EmailArchiving EmailArchiver
        {
            get; private set;
        }


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
                result = new LicenceValidationHelper(this.Log, Properties.Settings.Default.PublicKey, Properties.Settings.Default.LicenceKey).Validate();
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

                Thread background = new Thread(() => Run());
                background.Name = "Background";
                background.Start();
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

            StartLogging();

            synchronisationContext = new SyncContext(outlookApp);
            contactSynchroniser = new ContactSyncing("CS", synchronisationContext);
            taskSynchroniser = new TaskSyncing("TS", synchronisationContext);
            appointmentSynchroniser = new AppointmentSyncing("AS", synchronisationContext);
            EmailArchiver = new EmailArchiving("EM", synchronisationContext.Log);

            var outlookExplorer = outlookApp.ActiveExplorer();
            this.objExplorer = outlookExplorer;
            outlookExplorer.FolderSwitch -= objExplorer_FolderSwitch;
            outlookExplorer.FolderSwitch += objExplorer_FolderSwitch;

            // TODO: install/remove these event handlers when settings.AutoArchive changes:
            outlookApp.NewMailEx += new Outlook.ApplicationEvents_11_NewMailExEventHandler(this.Application_NewMail);
            outlookApp.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(this.Application_ItemSend);

            ((Outlook.ApplicationEvents_11_Event)Application).Quit += new Outlook.ApplicationEvents_11_QuitEventHandler(ThisAddIn_Quit);

            if (OutlookVersion < OutlookMajorVersion.Outlook2010)
            {
                outlookApp.ItemContextMenuDisplay += new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(this.Application_ItemContextMenuDisplay);
                var menuBar = outlookExplorer.CommandBars.ActiveMenuBar;
                objSuiteCRMMenuBar2007 = (Office.CommandBarPopup)menuBar.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, true);
                if (objSuiteCRMMenuBar2007 != null)
                {
                    ConstructOutlook2007MenuBar();
                }
            }
            else
            {
                //For Outlook version 2010 and greater
                //var app = this.Application;
                //app.FolderContextMenuDisplay += new Outlook.ApplicationEvents_11_FolderContextMenuDisplayEventHander(this.app_FolderContextMenuDisplay);
            }
        }

        private void ConstructOutlook2007MenuBar()
        {
            objSuiteCRMMenuBar2007.Caption = "SuiteCRM";
            this.btnArchive = (Office.CommandBarButton)this.objSuiteCRMMenuBar2007.Controls.Add(Office.MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            this.btnArchive.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
            this.btnArchive.Caption = "Archive";
            this.btnArchive.Picture = RibbonImageHelper.Convert(Resources.SuiteCRMLogo);
            this.btnArchive.Click += new Office._CommandBarButtonEvents_ClickEventHandler(this.cbtnArchive_Click);
            this.btnArchive.Visible = true;
            this.btnArchive.BeginGroup = true;
            this.btnArchive.TooltipText = "Archive selected emails to SuiteCRM";
            this.btnArchive.Enabled = true;
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

        /// <summary>
        /// Check the licence; if valid, try to login; if successful, do normal processing. If either 
        /// check fails, give the user the options of reconfiguring or disabling the add-in.
        /// </summary>
        private void Run()
        {
            bool success = false, disable = false;
            for (success = false; !(success || disable);)
            {
                success = this.VerifyLicenceKey();

                if (success)
                {
                    log.Info("Licence verified...");

                    success = this.SuiteCRMAuthenticate();

                    if (success)
                    {
                        log.Info("Authentication succeeded...");
                    }
                    else
                    {
                        disable = this.ShowReconfigureOrDisable("Login to CRM failed");
                        Log.Info("User chose to disable add-in after licence check succeeded but login to CRM failed.");
                    }
                }
                else
                {
                    disable = this.ShowReconfigureOrDisable("Licence check failed");
                    Log.Info("User chose to disable add-in after licence check failed.");
                }
            }

            if (success && !disable)
            {
                log.Info("Starting normal operations.");

                DaemonWorker.Instance.AddTask(new FetchEmailCategoriesAction());
                StartSynchronisationProcesses();
                this.IsLicensed = true;
            }
            else if (disable)
            {
                log.Info("Disabling addin at user request");
                this.Disable();
            }
            else
            {
                /* it's possible for both success AND disable to be true (if login to CRM fails); 
                 * but logically if success is false disabel must be true, so this branch should
                 * never be reached. */
                log.Error($"In ThisAddIn.Run: success is {success}; disable is {disable}; impossible state, disabling.");
            }
        }

        private void Disable()
        {
            Log.Warn("Disabling add-in");
            int i = 0;

            foreach (COMAddIn addin in Application.COMAddIns)
            {
                if (ProgId.Equals(addin.ProgId))
                {
                    Log.Debug($"ThisAddIn.Disable: Disabling instance {++i} of addin {ProgId}");
                    addin.Connect = false;
                }
                else
                {
                    Log.Debug($"ThisAddIn.Disable: Ignoring addin {addin.ProgId}");
                }
            }

            // Application.COMAddIns.Item(ProgId).Connect = false;
        }

        /// <summary>
        /// Show the reconfigure or disable dialogue with this summary of the problem.
        /// </summary>
        /// <param name="summary">A summary of the problem that caused the dialogue to be shown.</param>
        /// <returns>true if the user chose to disable the add-in.</returns>
        private bool ShowReconfigureOrDisable(string summary)
        {
            bool result;

            switch (new ReconfigureOrDisableDialog(summary).ShowDialog())
            {
                case DialogResult.OK:
                    /* if licence key does not validate, show the settings form to allow the user to enter
                     * a (new) key, and retry. */
                    Log.Info("User chose to reconfigure add-in");
                    this.ShowSettingsForm();
                    result = false;
                    break;
                case DialogResult.Cancel:
                    Log.Info("User chose to disable add-in");
                    result = true;
                    break;
                default:
                    log.Warn("Unexpected response from ReconfigureOrDisableDialog");
                    result = true;
                    break;
            }

            return result;
        }

        public static string LogDirPath =>
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
            "\\SuiteCRMOutlookAddIn\\Logs\\";

        private void StartLogging()
        {
            log = Log4NetLogger.FromFilePath("add-in", LogDirPath + "suitecrmoutlook.log", () => GetLogHeader(), Properties.Settings.Default.LogLevel);
            clsSuiteCRMHelper.SetLog(log);
        }

        private void LogKeySettings()
        {
            foreach (var s in GetKeySettings())
            {
                log.Error(s);
            }
        }

        private IEnumerable<string> GetLogHeader()
        {
            List<string> result = new List<string>();

            try
            {
                result.Add($"{AddInTitle} v{AddInVersion} in Outlook version {this.Application.Version}");
                result.AddRange(GetKeySettings());
            }
            catch (Exception any)
            {
                result.Add($"Exception {any.GetType().Name} '{any.Message}' while printing log header");
            }

            return result;
        }

        private IEnumerable<string> GetKeySettings()
        {
            yield return "Auto-archiving: " + (Properties.Settings.Default.AutoArchive ? "ON" : "off");
            yield return $"Logging level: {Properties.Settings.Default.LogLevel}";
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
            DoOrLogError(() => this.EmailArchiver.Start(), "Starting email archiver");
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

        public void ShowAddressBook()
        {
            if (HasCrmUserSession && this.IsLicensed)
            {
                new frmAddressBook().Show();
            }
            else
            {
                ReconfigureOrDisable();
            }
        }

        public void ShowSettingsForm()
        {
            var settingsForm = new SettingsDialog();
            settingsForm.SettingsChanged += (sender, args) => this.LogKeySettings();
            settingsForm.ShowDialog();
        }

        public void ShowArchiveForm()
        {
            ArchiveDialog objForm = new ArchiveDialog();
            objForm.ShowDialog();
        }

        internal void ManualArchive()
        {
            if (HasCrmUserSession && IsLicensed)
            {
                ShowArchiveForm();
            }
            else
            {
                ReconfigureOrDisable();
            }
        }

        private void ReconfigureOrDisable()
        {
            if (!HasCrmUserSession)
            {
                if (this.ShowReconfigureOrDisable("Login to CRM failed"))
                {
                    this.Disable();
                }
            }
            else if (!IsLicensed)
            {
                if (this.ShowReconfigureOrDisable("Licence check failed"))
                {
                    this.Disable();
                }
            }
        }

        /// <summary>
        /// Handle the quit signal.
        /// </summary>
        private void ThisAddIn_Quit()
        {
            Log.Info("ThisAddIn_Quit: signalled to quit");

            this.ShutdownAll();

            log.Info("ThisAddIn_Quit: shutdown complete.");
        }

        /// <summary>
        /// Handle the shutdown signal; this doesn't seem to be being received, but 
        /// may be in some versions of Outlook.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Log.Info("ThisAddIn_Shutdown: shutting down normally");

            ShutdownAll();

            log.Info("ThisAddIn_Shutdown: shutdown complete.");
        }

        /// <summary>
        /// Everything needed to shut down.
        /// </summary>
        private void ShutdownAll()
        {
            try
            {
                if (this.CommandBarExists("SuiteCRM"))
                {
                    Log.Info("ThisAddIn_Shutdown: Removing SuiteCRM command bar");
                    this.objSuiteCRMMenuBar2007.Delete();
                }

                this.UnregisterEvents();
                this.ShutdownProcesses();

                if (SuiteCRMUserSession != null)
                {
                    SuiteCRMUserSession.LogOut();
                }
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.ThisAddIn_Shutdown", ex);
            }
        }


        /// <summary>
        /// Shutdown all running RepeatingProcess instances, showing a progress bar 
        /// dialog if this cannot be done immediately, 
        /// </summary>
        private void ShutdownProcesses()
        {
            int stillToDo = RepeatingProcess.PrepareShutdownAll(this.log);

            if (stillToDo != 0)
            {
                new ShuttingDownDialog(stillToDo, this.log).ShowDialog();
            }

            log.Debug("ShutdownProcesses: complete");
        }

        private void UnregisterEvents()
        {
            try
            {
                Log.Info("ThisAddIn.UnregisterEvents: Removing context menu display event handler");
                this.Application.ItemContextMenuDisplay -= new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(this.Application_ItemContextMenuDisplay);
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.UnregisterEvents", ex);
            }
            try
            {
                Log.Info("ThisAddIn.UnregisterEvents: Removing archive button click event handler");
                this.btnArchive.Click -= new Office._CommandBarButtonEvents_ClickEventHandler(this.cbtnArchive_Click);
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.UnregisterEvents", ex);
            }

            try
            {
                Log.Info("ThisAddIn.UnregisterEvents: Removing new mail event handler");

                Outlook.ApplicationEvents_11_NewMailExEventHandler handler = new Outlook.ApplicationEvents_11_NewMailExEventHandler(this.Application_NewMail);

                if (handler != null)
                {
                    this.objExplorer.Application.NewMailEx -= handler;
                }
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.UnregisterEvents", ex);
            }

            try
            {
                Log.Info("ThisAddIn.UnregisterEvents: Removing archive item send event handler");
                this.objExplorer.Application.ItemSend -= new Outlook.ApplicationEvents_11_ItemSendEventHandler(this.Application_ItemSend);
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.UnregisterEvents", ex);
            }

            DisposeOf(appointmentSynchroniser);
            DisposeOf(contactSynchroniser);
            DisposeOf(taskSynchroniser);
        }

        /// <summary>
        /// Dispose of this disposable object, with extra logging.
        /// </summary>
        /// <param name="toDispose">The object of which to dispose.</param>
        private void DisposeOf(IDisposable toDispose)
        {
            if (toDispose != null)
            {
                try
                {
                    Log.Debug($"ThisAddIn.UnregisterEvents: Disposing of {toDispose.GetType().Name}");
                    toDispose.Dispose();
                }
                catch (Exception ex)
                {
                    log.Error($"DisposeOfSynchroniser: Failed to dispose of instance of {toDispose.GetType().Name}", ex);
                }
            }
            else
            {
                log.Error("Attempt to dispose of null reference?");
            }
        }

        private bool CommandBarExists(string name)
        {
            try
            {
                var explorer = Application.ActiveExplorer();
                return (explorer != null && explorer.CommandBars[name] != null);
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
                objMainMenu.Picture = RibbonImageHelper.Convert(Resources.SuiteCRMLogo);
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
                if (this.IsLicensed && Properties.Settings.Default.AutoArchive)
                {
                    ProcessNewMailItem(EmailArchiveType.Sent, item as Outlook.MailItem);
                }
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
                if (this.IsLicensed && Properties.Settings.Default.AutoArchive)
                {
                    ProcessNewMailItem(
                        EmailArchiveType.Inbound,
                        Application.Session.GetItemFromID(EntryID) as Outlook.MailItem);
                }
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
            else
            {
                this.EmailArchiver.ProcessEligibleNewMailItem(mailItem, archiveType);
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

        public bool SuiteCRMAuthenticate()
        {
            return HasCrmUserSession ? true : Authenticate();
        }

        public bool Authenticate()
        {
            bool result = false;
            try
            {
                if (Properties.Settings.Default.Host != String.Empty)
                {
                    SuiteCRMUserSession = 
                        new SuiteCRMClient.UserSession(
                            Properties.Settings.Default.Host,
                            Properties.Settings.Default.Username,
                            Properties.Settings.Default.Password, 
                            Properties.Settings.Default.LDAPKey, 
                            log,
                            Properties.Settings.Default.RestTimeout);
                    SuiteCRMUserSession.AwaitingAuthentication = true;
                    try
                    {
                        if (Properties.Settings.Default.IsLDAPAuthentication)
                        {
                            SuiteCRMUserSession.AuthenticateLDAP();
                        }
                        else
                        {
                            SuiteCRMUserSession.Login();
                        }

                        if (SuiteCRMUserSession.IsLoggedIn)
                        {
                            result = true;
                        }
                    }
                    catch (Exception any)
                    {
                        ShowAndLogError(any, "Failure while trying to authenticate to CRM", "Login failure");
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
                            Properties.Settings.Default.RestTimeout);
                }

                SuiteCRMUserSession.AwaitingAuthentication = false;
            }
            catch (Exception ex)
            {
                log.Error("ThisAddIn.Authenticate", ex);
            }

            return result;
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

        /// <summary>
        /// True if this is a licensed copy of the add-in.
        /// </summary>
        public bool IsLicensed { get; private set; } = false;

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

        public void ShowAndLogError(Exception error, string message, string summary)
        {
            MessageBox.Show(
                message,
                summary,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            Log.Error(summary, error);
        }
    }
}
