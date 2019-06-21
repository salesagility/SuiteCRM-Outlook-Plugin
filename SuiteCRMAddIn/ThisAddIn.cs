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
    using Extensions;
    using Helpers;
    using Microsoft.Office.Core;
    using NGettext;
    using Properties;
    using SuiteCRMClient;
    using SuiteCRMClient.Email;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Text;
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
        private ContactSynchroniser contactSynchroniser;
        private TaskSynchroniser taskSynchroniser;
        private MeetingsSynchroniser meetingSynchroniser;
        private CallsSynchroniser callSynchroniser;

        /// <summary>
        /// #2246: Discriminate between calls and meetings when adding and updating.
        /// </summary>
        internal CallsSynchroniser CallsSynchroniser => this.callSynchroniser;
        /// <summary>
        /// #2246: Discriminate between calls and meetings when adding and updating.
        /// </summary>
        internal MeetingsSynchroniser MeetingsSynchroniser => this.meetingSynchroniser;

        internal ContactSynchroniser ContactsSynchroniser => this.contactSynchroniser;

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
            }
            catch (System.Configuration.SettingsPropertyNotFoundException ex)
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
            /* we need logging before we start the daemon */
            StartLogging();

            DaemonWorker.Instance.AddTask(new DeferredStartupAction());
        }

        /// <summary>
        ///  Actually perform the startup of all the addin's services.
        /// </summary>
        /// <remarks>
        /// Called by <see cref="Daemon.DeferredStartupAction"/>, q.v.
        /// </remarks>
        internal void DeferredStartup()
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

            this.MaybeUpgradeSettings();

            OutlookVersion = (OutlookMajorVersion)Convert.ToInt32(outlookApp.Version.Split('.')[0]);

            synchronisationContext = new SyncContext(outlookApp);
            callSynchroniser = new CallsSynchroniser("AS", synchronisationContext);
            contactSynchroniser = new ContactSynchroniser("CS", synchronisationContext);
            meetingSynchroniser = new MeetingsSynchroniser("MS", synchronisationContext);
            taskSynchroniser = new TaskSynchroniser("TS", synchronisationContext);
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
        }

        internal void ManualSyncContact()
        {
            new ManualSyncContactForm(string.Join("; ", SelectedContacts.Select(x => $"{x.FullName}, {x.Email1Address}"))).ShowDialog();
        }

        /// <summary>
        /// Check whether an upgrade of settings is required, and if so, do it.
        /// </summary>
        private void MaybeUpgradeSettings()
        {
            /* Attempt to fix 'settings getting wiped' bug #187;
             * see https://stackoverflow.com/questions/2201819/why-are-persisted-user-settings-not-loaded
             */
            if (Settings.Default.NeedsUpgrade)
            {
                Settings.Default.Upgrade();
                Settings.Default.NeedsUpgrade = false;
                Settings.Default.Save();
            }
            /* The above doesn't always work; Settings.Default.NeedsUpgrade doesn't
             * always fire when it should. I hypothesise that this is because we are
             * an Add-in, not a first class application, and our settings directory
             * actually tracks the Microsoft Office version. So check if a key setting
             * is missing, and if it is, attempt upgrade.
             */
            else if (String.IsNullOrWhiteSpace(Settings.Default.LicenceKey))
            {
                try
                {
                    Log.Debug("No licence key? Trying upgrade...");
                    Settings.Default.Upgrade();
                    Settings.Default.Save();
                    Log.Debug("Upgrade succeeded.");
                }
                catch (Exception any)
                {
                    /* This will fail in the case of a new installation, but that's OK. 
                     * However, log it, in case it also fails at other times. */
                    ErrorHandler.Handle("Failure while attempting to upgrade settings.", any);
                }
            }
        }

        private void ConstructOutlook2007MenuBar()
        {
            objSuiteCRMMenuBar2007.Caption = "SuiteCRM";
            this.btnArchive = (Office.CommandBarButton)this.objSuiteCRMMenuBar2007.Controls.Add(Office.MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
            this.btnArchive.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
            this.btnArchive.Caption = catalogue.GetString("Archive");
            this.btnArchive.Picture = RibbonImageHelper.Convert(Resources.SuiteCRMLogo);
            this.btnArchive.Enabled &= this.HasCrmUserSession;
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
                        disable = this.ShowReconfigureOrDisable(catalogue.GetString("Login to CRM failed"), true) == DialogResult.Cancel;
                    }
                }
                else
                {
                    disable = this.ShowReconfigureOrDisable(catalogue.GetString("Licence check failed"), true) == DialogResult.Cancel;
                }
            }

            if (success && !disable)
            {
                log.Info(catalogue.GetString("Starting normal operations."));

                DaemonWorker.Instance.AddTask(new FetchEmailCategoriesAction());
                StartConfiguredSynchronisationProcesses();
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
                Log.Error(
                    catalogue.GetString(
                        "In {0}: success is {1}; disable is {2}; impossible state, disabling.",
                        "ThisAddIn.Run",
                        success,
                        disable));
                this.Disable();
            }
        }

        internal void Disable()
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
        /// <param name="allowRetry">if true, the action may be retried.</param>
        /// <returns>true if the user chose to disable the add-in.</returns>
        internal DialogResult ShowReconfigureOrDisable(string summary, bool allowRetry = false)
        {
            DialogResult result;

            result = new ReconfigureOrDisableDialog(summary, allowRetry).ShowDialog();
            switch (result)
            {
                case DialogResult.Retry:
                    Log.Info(catalogue.GetString("User chose to retry connection"));
                    break;
                case DialogResult.OK:
                    /* if licence key does not validate, show the settings form to allow the user to enter
                     * a (new) key, and retry. */
                    Log.Info(catalogue.GetString("User chose to reconfigure add-in"));
                    this.ShowSettingsForm();
                    break;
                case DialogResult.Cancel:
                    Log.Info(catalogue.GetString("User chose to disable add-in"));
                    break;
                default:
                    log.Warn(catalogue.GetString("Unexpected response from ReconfigureOrDisableDialog"));
                    result = DialogResult.OK;
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
                log.Debug(s);
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
                ErrorHandler.Handle($"Failure while attempting to change folder to {this.objExplorer.CurrentFolder.FolderPath}", ex);
            }
        }

        /// <summary>
        /// Start all synchronisation processes that are configured, if they 
        /// are not already running.
        /// </summary>
        public void StartConfiguredSynchronisationProcesses()
        {
            StartSynchroniserIfConfigured(this.callSynchroniser);
            StartSynchroniserIfConfigured(this.contactSynchroniser);
            StartSynchroniserIfConfigured(this.meetingSynchroniser);
            StartSynchroniserIfConfigured(this.taskSynchroniser);
            StartSynchroniserIfConfigured(this.EmailArchiver);
        }

        /// <summary>
        /// Start all synchronisation processes that are not configured to run,
        /// if they are already running.
        /// </summary>
        public void StopUnconfiguredSynchronisationProcesses()
        {
            StopSynchroniserIfUnconfigured(this.callSynchroniser);
            StopSynchroniserIfUnconfigured(this.contactSynchroniser);
            StopSynchroniserIfUnconfigured(this.meetingSynchroniser);
            StopSynchroniserIfUnconfigured(this.taskSynchroniser);
            StopSynchroniserIfUnconfigured(this.EmailArchiver);
        }

        /// <summary>
        /// Start this synchronisation process, if it is configured to run, 
        /// provided it is not already running.
        /// </summary>
        /// <param name="synchroniser">The synchroniser to start.</param>
        private void StartSynchroniserIfConfigured(Synchroniser synchroniser)
        {
            if (synchroniser != null &&
                synchroniser.Direction != SyncDirection.Direction.Neither &&
                !synchroniser.IsActive)
            {
                DoOrLogError(() =>
                    synchroniser.Start(),
                    catalogue.GetString("Starting {0}", new object[] { synchroniser.GetType().Name }));
            }
        }


        /// <summary>
        /// Start this archiver process, if it is configured to run, 
        /// provided it is not already running.
        /// </summary>
        /// <param name="archiver">The archiver to start.</param>
        private void StartSynchroniserIfConfigured(EmailArchiving archiver)
        {
            if (archiver != null && Settings.Default.AutoArchive == true && !archiver.IsActive)
            {
                DoOrLogError(() =>
                    archiver.Start(),
                    catalogue.GetString("Starting {0}", new object[] { archiver.GetType().Name }));
            }
        }


        /// <summary>
        /// Stop this synchroniser if it is active and is configured to be inactive.
        /// </summary>
        /// <param name="synchroniser">The synchroniser to stop.</param>
        private void StopSynchroniserIfUnconfigured(Synchroniser synchroniser)
        {
            if (synchroniser != null &&
                synchroniser.Direction == SyncDirection.Direction.Neither &&
                synchroniser.IsActive)
            {
                synchroniser.Stop();
            }
        }


        /// <summary>
        /// Stop this archiver if it is active and is configured to be inactive.
        /// </summary>
        /// <param name="archiver">The archiver to stop.</param>
        private void StopSynchroniserIfUnconfigured(EmailArchiving archiver)
        {
            if (archiver != null && Settings.Default.AutoArchive == false && archiver.IsActive)
            {
                archiver.Stop();
            }
        }


        private void cbtnArchive_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ShowArchiveForm();
        }

        private void cbtnSettings_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            DoOrLogError(() =>
                ShowSettingsForm());
        }

        public DialogResult ShowSettingsForm()
        {
            var settingsForm = new SettingsDialog();
            settingsForm.SettingsChanged += (sender, args) => this.LogKeySettings();
            return settingsForm.ShowDialog();
        }

        public void ShowAddressBook()
        {
            if (this.HasCrmUserSession)
            {
                DoOrLogError(() =>
                new frmAddressBook().ShowDialog());
            }
            else
            {
                MessageBox.Show("Please wait: SuiteSRM AddIn has not yet completed connections", 
                    "Waiting for connection", 
                    MessageBoxButtons.OK);
            }
        }

        public void ShowArchiveForm()
        {
            if (this.HasCrmUserSession)
            {
                if (this.SelectedEmails.Any())
            {
                DoOrLogError( () => 
                new ArchiveDialog(this.SelectedEmails, EmailArchiveReason.Manual).ShowDialog());
            }
            }
            else
            {
                MessageBox.Show("Please wait: SuiteSRM AddIn has not yet completed connections",
                    "Waiting for connection",
                    MessageBoxButtons.OK);
            }
        }

        internal void ManualArchive()
        {
            if (HasCrmUserSession)
            {
                ShowArchiveForm();
            }
            else
            {
                MessageBox.Show("Please wait: SuiteSRM AddIn has not yet completed connections",
                    "Waiting for connection",
                    MessageBoxButtons.OK);
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
                SyncStateManager.Instance.BruteForceSaveAll();
                this.ShutdownProcesses();

                if (SuiteCRMUserSession != null)
                {
                    SuiteCRMUserSession.LogOut();
                }

                DisposeOf(callSynchroniser);
                DisposeOf(contactSynchroniser);
                DisposeOf(meetingSynchroniser);
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
                ErrorHandler.Handle(catalogue.GetString("Failed while trying to handle a sent email"), ex);
            }
        }


        private void Application_NewMail(string EntryID)
        {
            log.Debug(catalogue.GetString("Outlook NewMail: email received event"));
            try
            {
                if (this.IsLicensed)
                {
                    var item = Application.Session.GetItemFromID(EntryID);

                    if (item is Outlook.MailItem && Properties.Settings.Default.AutoArchive)
                    {
                        ProcessNewMailItem(EmailArchiveReason.Inbound,
                                            item as Outlook.MailItem,
                                            Settings.Default.ExcludedEmails);
                    }
                    else if (item is Outlook.MeetingItem && SyncDirection.AllowOutbound(Properties.Settings.Default.SyncMeetings))
                    {
                        ProcessNewMeetingItem(item as Outlook.MeetingItem);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle(catalogue.GetString("Failed while trying to handle a received email"), ex);
            }
        }

        private void ProcessNewMeetingItem(Outlook.MeetingItem meetingItem)
        {
            string vCalId = meetingItem.GetVCalId();

            if (CrmId.IsValid(vCalId) && RestAPIWrapper.GetEntry(MeetingsSynchroniser.DefaultCrmModule, vCalId, new string[] { "id" }) != null)
            {
                meetingItem.GetAssociatedAppointment(false).SetCrmId(CrmId.Get(vCalId));
            }
        }

        private bool ProcessNewMailItem(EmailArchiveReason archiveType, Outlook.MailItem mailItem, string excludedEmails = "")
        {
            bool result;

            if (mailItem == null)
            {
                log.Info(catalogue.GetString("New 'mail item' was null"));
                result = false;
            }
            else
            {
                this.EmailArchiver.ProcessEligibleNewMailItem(mailItem, archiveType, excludedEmails);
                result = true;
            }

            return result;
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

        public bool SuiteCRMAuthenticate() => HasCrmUserSession ? true : Authenticate();

        /// <summary>
        /// Authenticate against CRM using parameters taken from default settings.
        /// </summary>
        /// <returns>true on success</returns>
        public bool Authenticate() {
            return Authenticate(Properties.Settings.Default.Host,
                    Properties.Settings.Default.Username,
                    Properties.Settings.Default.Password,
                    Properties.Settings.Default.LDAPKey);
        }

        /// <summary>
        /// Authenticate against CRM using these parameters.
        /// </summary>
        /// <param name="host"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="ldapKey"></param>
        /// <returns>True on success.</returns>
        public bool Authenticate(string host, string username, string password, string ldapKey)
        {
            bool result = false;
            try
            {
                if (Properties.Settings.Default.Host != String.Empty)
                {
                    ReinitialiseSession(host, username, password, ldapKey);
                    try
                    {
                        if (SuiteCRMUserSession.Login())
                        {
                            LogServerVersion();

                            result = true;
                        }
                    }
                    catch (Exception any)
                    {
                        Log.Error("Failure while trying to authenticate to CRM", any);
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
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.Authenticate", ex);
            }

            return result;
        }

        /// <summary>
        /// Replace the existing user session, if any, with a new one using 
        /// these parameters.
        /// </summary>
        public void ReinitialiseSession(string host, string username, string password, string ldapKey)
        {
            SuiteCRMUserSession = new SuiteCRMClient.UserSession(host, username, password, ldapKey,
                ThisAddIn.AddInTitle,
                log,
                Properties.Settings.Default.RestTimeout);
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
            else if (!string.IsNullOrWhiteSpace(info.SugarVersion))
            {
                log.Info($"Connected to an instance of Sugar version {info.SugarVersion}.");
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

        public IEnumerable<Outlook.ContactItem> SelectedContacts
        {
            get
            {
                var selection = Application.ActiveExplorer()?.Selection;
                if (selection == null) yield break;
                foreach (object e in selection)
                {
                    var contact = e as Outlook.ContactItem;
                    if (contact != null) yield return contact;
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

        /// <summary>
        /// Count how many items there are whose state I am monitoring.
        /// </summary>
        /// <returns>The number of items I am monitoring.</returns>
        internal int CountItems()
        {
            return SyncStateManager.Instance.CountItems();
        }


        /// <summary>
        /// Get all the synchronisable items I'm tracking - at present, only as WithRemovableSynchronisationProperties objects.
        /// </summary>
        /// <returns>all the synchronisable items I'm tracking</returns>
        internal IEnumerable<WithRemovableSynchronisationProperties> GetSynchronisableItems()
        {
            List<WithRemovableSynchronisationProperties> result = new List<WithRemovableSynchronisationProperties>();

            if (this.callSynchroniser != null)
            {
                result.AddRange(this.callSynchroniser.GetSynchronisedItems());
            }
            if (this.contactSynchroniser != null)
            {
                result.AddRange(this.contactSynchroniser.GetSynchronisedItems());
            }
            if (this.meetingSynchroniser != null)
            {
                result.AddRange(this.meetingSynchroniser.GetSynchronisedItems());
            }
            if (this.taskSynchroniser != null)
            {
                result.AddRange(this.taskSynchroniser.GetSynchronisedItems());
            }

            return result;
        }
    }
}
