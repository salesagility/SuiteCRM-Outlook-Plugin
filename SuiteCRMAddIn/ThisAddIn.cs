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
    using NGettext;
    using SuiteCRMAddIn.Properties;
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
        public const string ProgId = "SuiteCRMAddIn";
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

        /// <summary>
        /// Internationalisation (118n) strings dictionary
        /// </summary>
        public ICatalog catalogue = new Catalog(ProgId, "./Locale");

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
                this.log.Error(catalogue.GetString("Licence key was not yet set up"), ex);
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

            /* Attempt to fix 'settings getting wiped' bug 187;
             * see https://stackoverflow.com/questions/2201819/why-are-persisted-user-settings-not-loaded
             */ 
            if (Settings.Default.NeedsUpgrade)
            {
                Settings.Default.Upgrade();
                Settings.Default.NeedsUpgrade = false;
                Settings.Default.Save();
            }

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
            this.btnArchive.Caption = catalogue.GetString("Archive");
            this.btnArchive.Picture = RibbonImageHelper.Convert(Resources.SuiteCRMLogo);
            this.btnArchive.Click += new Office._CommandBarButtonEvents_ClickEventHandler(this.cbtnArchive_Click);
            this.btnArchive.Visible = true;
            this.btnArchive.BeginGroup = true;
            this.btnArchive.TooltipText = catalogue.GetString("Archive selected emails to SuiteCRM");
            this.btnArchive.Enabled = true;
            this.btnSettings = (Office.CommandBarButton)this.objSuiteCRMMenuBar2007.Controls.Add(Office.MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            this.btnSettings.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
            this.btnSettings.Caption = catalogue.GetString("Settings");
            this.btnSettings.Click += new Office._CommandBarButtonEvents_ClickEventHandler(this.cbtnSettings_Click);
            this.btnSettings.Visible = true;
            this.btnSettings.BeginGroup = true;
            this.btnSettings.TooltipText = catalogue.GetString("SuiteCRM Settings");
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
                    log.Info(catalogue.GetString("Licence verified..."));

                    success = this.SuiteCRMAuthenticate();

                    if (success)
                    {
                        log.Info(catalogue.GetString("Authentication succeeded..."));
                    }
                    else
                    {
                        disable = this.ShowReconfigureOrDisable(catalogue.GetString("Login to CRM failed"));
                        Log.Info(catalogue.GetString("User chose to disable add-in after licence check succeeded but login to CRM failed."));
                    }
                }
                else
                {
                    disable = this.ShowReconfigureOrDisable(catalogue.GetString("Licence check failed"));
                    Log.Info(catalogue.GetString("User chose to disable add-in after licence check failed."));
                }
            }

            if (success && !disable)
            {
                log.Info(catalogue.GetString("Starting normal operations."));

                DaemonWorker.Instance.AddTask(new FetchEmailCategoriesAction());
                StartSynchronisationProcesses();
                this.IsLicensed = true;
            }
            else if (disable)
            {
                log.Info(catalogue.GetString("Disabling addin at user request"));
                this.Disable();
            }
            else
            {
                /* it's possible for both success AND disable to be true (if login to CRM fails); 
                 * but logically if success is false disable must be true, so this branch should
                 * never be reached. */
                log.Error(
                    catalogue.GetString(
                        "In {0}: success is {1}; disable is {2}; impossible state, disabling.",
                        "ThisAddIn.Run",
                        success, 
                        disable));
                this.Disable();
            }
        }

        private void Disable()
        {
            const string methodName = "ThisAddIn.Disable";
            Log.Warn(catalogue.GetString("Disabling add-in"));
            int i = 0;

            foreach (COMAddIn addin in Application.COMAddIns)
            {
                if (ProgId.Equals(addin.ProgId))
                {
                    Log.Debug(catalogue.GetString("{0}: Disabling instance {1} of addin {2}", methodName, ++i, ProgId));
                    addin.Connect = false;
                }
                else
                {
                    Log.Debug(catalogue.GetString("{0}: Ignoring addin {1}", methodName, addin.ProgId));
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
                    Log.Info(catalogue.GetString("User chose to reconfigure add-in"));
                    this.ShowSettingsForm();
                    result = false;
                    break;
                case DialogResult.Cancel:
                    Log.Info(catalogue.GetString("User chose to disable add-in"));
                    result = true;
                    break;
                default:
                    log.Warn(catalogue.GetString("Unexpected response from ReconfigureOrDisableDialog"));
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
            RestAPIWrapper.SetLog(log);
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
                result.Add(catalogue.GetString("{0} v{1} in Outlook version {2}", AddInTitle, AddInVersion, this.Application.Version));
                result.AddRange(GetKeySettings());
            }
            catch (Exception any)
            {
                result.Add(catalogue.GetString("Exception {0} '{1}' while printing log header", any.GetType().Name, any.Message));
            }

            return result;
        }

        private IEnumerable<string> GetKeySettings()
        {
            yield return catalogue.GetString("Auto-archiving: ") + 
                (Settings.Default.AutoArchive ? catalogue.GetString("ON") : catalogue.GetString("off"));
            yield return catalogue.GetString("Logging level: {0}", Settings.Default.LogLevel);
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
            DoOrLogError(() => this.appointmentSynchroniser.Start(), catalogue.GetString("Starting appointments synchroniser"));
            DoOrLogError(() => this.contactSynchroniser.Start(), catalogue.GetString("Starting contacts synchroniser"));
            DoOrLogError(() => this.taskSynchroniser.Start(), catalogue.GetString("Starting tasks synchroniser"));
            DoOrLogError(() => this.EmailArchiver.Start(), catalogue.GetString("Starting email archiver"));
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
                if (this.ShowReconfigureOrDisable(catalogue.GetString("Login to CRM failed")))
                {
                    this.Disable();
                }
            }
            else if (!IsLicensed)
            {
                if (this.ShowReconfigureOrDisable(catalogue.GetString("Licence check failed")))
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
            const string methodName = "ThisAddIn_Quit";

            Log.Info(catalogue.GetString("{0}: signalled to quit.", methodName));

            this.ShutdownAll();

            log.Info(catalogue.GetString("{0}: shutdown complete.", methodName));
        }

        /// <summary>
        /// Handle the shutdown signal; this doesn't seem to be being received, but 
        /// may be in some versions of Outlook.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            const string methodName = "ThisAddIn_Shutdown";
            Log.Info(catalogue.GetString("{0}: shutting down normally.", methodName));

            ShutdownAll();

            log.Info(catalogue.GetString("{0}: shutdown complete.", methodName));
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
                    Log.Info(catalogue.GetString("{0}: Removing SuiteCRM command bar", "ThisAddIn_ShutdownAll"));
                    this.objSuiteCRMMenuBar2007.Delete();
                }

                this.UnregisterEvents();
                this.ShutdownProcesses();

                if (SuiteCRMUserSession != null)
                {
                    SuiteCRMUserSession.LogOut();
                }

                DisposeOf(appointmentSynchroniser);
                DisposeOf(contactSynchroniser);
                DisposeOf(taskSynchroniser);
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

            log.Debug(catalogue.GetString("ShutdownProcesses: complete"));
        }

        private void UnregisterEvents()
        {
            const string methodName = "ThisAddIn.UnregisterEvents";
            try
            {
                Log.Info(catalogue.GetString("{0}: Removing context menu display event handler", methodName));
                this.Application.ItemContextMenuDisplay -= 
                    new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(this.Application_ItemContextMenuDisplay);
            }
            catch (Exception ex)
            {
                log.Error(methodName, ex);
            }

            UnregisterButtonClickHandler(this.btnArchive, this.cbtnArchive_Click);
            UnregisterButtonClickHandler(this.btnSettings, this.cbtnSettings_Click);

            try
            {
                Log.Info(catalogue.GetString("{0}: Removing new mail event handler", methodName));

                Outlook.ApplicationEvents_11_NewMailExEventHandler handler = 
                    new Outlook.ApplicationEvents_11_NewMailExEventHandler(this.Application_NewMail);

                if (handler != null)
                {
                    this.objExplorer.Application.NewMailEx -= handler;
                }
            }
            catch (Exception ex)
            {
                log.Error(methodName, ex);
            }

            try
            {
                Log.Info(catalogue.GetString("{0}: Removing archive item send event handler", methodName));
                this.objExplorer.Application.ItemSend -= new Outlook.ApplicationEvents_11_ItemSendEventHandler(this.Application_ItemSend);
            }
            catch (Exception ex)
            {
                log.Error(methodName, ex);
            }
        }

        private void UnregisterButtonClickHandler(CommandBarButton button, _CommandBarButtonEvents_ClickEventHandler clickHandler)
        {
            string methodName = "ThisAddIn.UnregisterButtonClickHandler";
            if (button != null)
            {
                try
                {
                    Log.Info(catalogue.GetString("{0}: Removing archive button click event handler", methodName));
                    button.Click -= new Office._CommandBarButtonEvents_ClickEventHandler(clickHandler);
                }
                catch (Exception ex)
                {
                    log.Error(methodName, ex);
                }
            }
        }

        /// <summary>
        /// Dispose of this disposable object, with extra logging.
        /// </summary>
        /// <param name="toDispose">The object of which to dispose.</param>
        private void DisposeOf(IDisposable toDispose)
        {
            string methodName = "ThisAddIn.DisposeOf";

            if (toDispose != null)
            {
                try
                {
                    Log.Debug(catalogue.GetString("{0}: Disposing of {1}", methodName, toDispose.GetType().Name));
                    toDispose.Dispose();
                }
                catch (Exception ex)
                {
                    log.Error(catalogue.GetString("{0}: Failed to dispose of instance of {1}", methodName, toDispose.GetType().Name), ex);
                }
            }
            else
            {
                log.Error(catalogue.GetString("{0}: Attempt to dispose of null reference?", methodName));
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
                objMainMenu.Caption = catalogue.GetString("SuiteCRM Archive");
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
            log.Debug(catalogue.GetString("Outlook ItemSend: email sent event"));
            try
            {
                if (this.IsLicensed && Properties.Settings.Default.AutoArchive)
                {
                    ProcessNewMailItem(
                        EmailArchiveReason.Outbound, 
                        item as Outlook.MailItem,
                        Properties.Settings.Default.ExcludedEmails);
                }
            }
            catch (Exception ex)
            {
                log.Error(catalogue.GetString("ThisAddIn.Application_ItemSend"), ex);
            }
        }

        private void Application_NewMail(string EntryID)
        {
            log.Debug(catalogue.GetString("Outlook NewMail: email received event"));
            try
            {
                if (this.IsLicensed && Properties.Settings.Default.AutoArchive)
                {
                    ProcessNewMailItem(
                        EmailArchiveReason.Inbound,
                        Application.Session.GetItemFromID(EntryID) as Outlook.MailItem,
                        Properties.Settings.Default.ExcludedEmails);
                }
            }
            catch (Exception ex)
            {
                log.Error(catalogue.GetString("ThisAddIn.Application_NewMail"), ex);
            }
        }

        private void ProcessNewMailItem(EmailArchiveReason archiveType, Outlook.MailItem mailItem, string excludedEmails = "")
        {
            if (mailItem == null)
            {
                log.Info(catalogue.GetString("New 'mail item' was null"));
                return;
            }
            else
            {
                this.EmailArchiver.ProcessEligibleNewMailItem(mailItem, archiveType, excludedEmails);
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
                            ThisAddIn.AddInTitle,
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
                            LogServerVersion();

                            result = true;
                        }
                    }
                    catch (Exception any)
                    {
                        ShowAndLogError(
                            any, 
                            catalogue.GetString("Failure while trying to authenticate to CRM"), 
                            catalogue.GetString("Login failure"));
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
                            ThisAddIn.AddInTitle,
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

        /// <summary>
        /// Obtain the server version from the server, if specified, and write it to the log.
        /// </summary>
        private void LogServerVersion()
        {
            ServerInfo info = RestAPIWrapper.GetServerInfo();

            if (!string.IsNullOrWhiteSpace(info.SuiteCRMVersion))
            {
                log.Info($"Connected to an instance of SuiteCRM version {info.SuiteCRMVersion}.");
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
