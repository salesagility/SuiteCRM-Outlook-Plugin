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
namespace SuiteCRMAddIn.Dialogs
{
    using BusinessLogic;
    using Microsoft.Office.Interop.Outlook;
    using Microsoft.Win32;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Windows.Forms;
    using Exception = System.Exception;

    public partial class SettingsDialog : Form
    {
        public EventHandler SettingsChanged;

        public SettingsDialog()
        {
            InitializeComponent();
        }

        private ILogger Log => Globals.ThisAddIn.Log;

        private Microsoft.Office.Interop.Outlook.Application Application => Globals.ThisAddIn.Application;

        /// <summary>
        /// The CRM URL value at the time the dialog was invoked.
        /// </summary>
        private string oldUrl = Properties.Settings.Default.Host;

        private bool ValidateDetails()
        {
            if (SafelyGetText(txtURL) == string.Empty)
            {
                MessageBox.Show("Please enter a valid SuiteCRM URL");
                txtURL.Focus();
                return false;
            }

            if (SafelyGetText(txtUsername) == string.Empty)
            {
                MessageBox.Show("Please enter a valid SuiteCRM Username");
                txtUsername.Focus();
                return false;
            }

            if (SafelyGetText(txtPassword) == string.Empty)
            {
                MessageBox.Show("Please enter a valid SuiteCRM Password");
                txtPassword.Focus();
                return false;
            }

            if (chkEnableLDAPAuthentication.Checked)
            {
                if (SafelyGetText(txtLDAPAuthenticationKey) == string.Empty)
                {
                    MessageBox.Show("Please enter a valid LDAP authentication key");
                    txtLDAPAuthenticationKey.Focus();
                    return false;
                }
            }

            return true;
        }

        private void frmSettings_Load(object sender, EventArgs e)
        {
            using (new WaitCursor(this))
            {
                try
                {
                    AddInTitleLabel.Text = ThisAddIn.AddInTitle;
                    AddInVersionLabel.Text = "Version " + ThisAddIn.AddInVersion;
                    dotNetVersionLabel.Text = ".Net version " + CheckFor45DotVersion(Get45or451FromRegistry());

                    if (Globals.ThisAddIn.SuiteCRMUserSession == null)
                    {
                        Globals.ThisAddIn.ReinitialiseSession(
                                string.Empty, string.Empty, string.Empty, string.Empty);
                    }

                    LoadSettings();
                    LinkToLogFileDir.Text = ThisAddIn.LogDirPath;
                }
                catch (Exception ex)
                {
                    ErrorHandler.Handle("Failed while loading the settings form", ex);
                    // Swallow exception!
                }
            }
        }

        private static int Get45or451FromRegistry()
        {
            int result;
            try
            {
                using (RegistryKey ndpKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                    .OpenSubKey("SOFTWARE\\Microsoft\\NET Framework Setup\\NDP\\v4\\Full\\"))
                {
                    result = Convert.ToInt32(ndpKey.GetValue("Release"));
                }
            }
            catch (Exception)
            {
                result = 0;
            }

            return result;
        }



// Checking the version using >= will enable forward compatibility,  
// however you should always compile your code on newer versions of 
// the framework to ensure your app works the same. 
        private static string CheckFor45DotVersion(int releaseKey)
        {
            if (releaseKey >= 461808)
            {
                return "4.7.2";
            }
            if (releaseKey >= 461308)
            {
                return "4.7.1";
            }
            if (releaseKey >= 460798)
            {
                return "4.7";
            }
            if (releaseKey >= 394802)
            {
                return "4.6.2";
            }
            if (releaseKey >= 394254)
            {
                return "4.6.1";
            }
            if (releaseKey >= 393295)
            {
                return "4.6";
            }
            if (releaseKey >= 393273)
            {
                return "4.6 RC";
            }
            if ((releaseKey >= 379893))
            {
                return "4.5.2";
            }
            if ((releaseKey >= 378675))
            {
                return "4.5.1";
            }
            if ((releaseKey >= 378389))
            {
                return "4.5";
            }
            // This line should never execute. A non-null release key should mean 
            // that 4.5 is installed. 
            return "No 4.5 version detected";
        }

        private void LoadSettings()
        {
            if (Properties.Settings.Default.Host != string.Empty)
            {
                txtURL.Text = Properties.Settings.Default.Host;
                txtUsername.Text = Properties.Settings.Default.Username;
                txtPassword.Text = Properties.Settings.Default.Password;
                licenceText.Text = Properties.Settings.Default.LicenceKey;
            }
            this.chkEnableLDAPAuthentication.Checked = Properties.Settings.Default.IsLDAPAuthentication;
            this.txtLDAPAuthenticationKey.Text = Properties.Settings.Default.LDAPKey;

            this.cbEmailAttachments.Checked = Properties.Settings.Default.ArchiveAttachments;
            this.checkBoxAutomaticSearch.Checked = Properties.Settings.Default.AutomaticSearch;
            this.cbShowCustomModules.Checked = Properties.Settings.Default.ShowCustomModules;
            this.txtSyncMaxRecords.Text = Properties.Settings.Default.SyncMaxRecords.ToString();
            this.checkBoxShowRightClick.Checked = Properties.Settings.Default.PopulateContextLookupList;
            GetAccountAutoArchivingSettings();

            txtAutoSync.Text = Properties.Settings.Default.ExcludedEmails;

            dtpAutoArchiveFrom.Value = DateTime.Now.AddDays(0 - Properties.Settings.Default.DaysOldEmailToAutoArchive);
            chkShowConfirmationMessageArchive.Checked = Properties.Settings.Default.ShowConfirmationMessageArchive;

            if (chkEnableLDAPAuthentication.Checked)
            {
                labelKey.Enabled = true;
                txtLDAPAuthenticationKey.Enabled = true;
            }
            else
            {
                labelKey.Enabled = false;
                txtLDAPAuthenticationKey.Enabled = false;
            }

            logLevelSelector.DataSource = Enum.GetValues(typeof(LogEntryType))
                .Cast<LogEntryType>()
                .Select(p => new { Key = (int)p, Value = p.ToString() })
                .OrderBy(o => o.Key)
                .ToList();
            logLevelSelector.DisplayMember = "Value";
            logLevelSelector.ValueMember = "Key";

            try
            {
                logLevelSelector.SelectedValue = Convert.ToInt32(Properties.Settings.Default.LogLevel);
            }
            catch (NullReferenceException)
            {
                logLevelSelector.SelectedValue = LogEntryType.Error;
            }

            showErrorsSelector.DataSource = Enum.GetValues(typeof(ErrorHandler.PopupWhen))
                .Cast<ErrorHandler.PopupWhen>()
                .Select(p => new { Key = (int)p, Value = p.ToString() })
                .OrderBy(o => o.Key)
                .ToList();
            showErrorsSelector.DisplayMember = "Value";
            showErrorsSelector.ValueMember = "Key";

            try
            {
                showErrorsSelector.SelectedValue = Convert.ToInt32(Properties.Settings.Default.ShowExceptions);
            }
            catch (NullReferenceException)
            {
                showErrorsSelector.SelectedValue = ErrorHandler.PopupWhen.EveryTime;
            }

            startupDeferralInput.Value = Properties.Settings.Default.StartupDeferral;

            crmIdValidationSelector.DataSource = Enum.GetValues(typeof(CrmIdValidationPolicy.Policy))
                .Cast<CrmIdValidationPolicy.Policy>()
                .Select(p => new { Key = (int)p, Value = p.ToString() })
                .OrderBy(o => o.Key)
                .ToList();
            crmIdValidationSelector.DisplayMember = "Value";
            crmIdValidationSelector.ValueMember = "Key";

            try
            {
                crmIdValidationSelector.SelectedValue = Convert.ToInt32(Properties.Settings.Default.CrmIdValidationPolicy);
            }
            catch (NullReferenceException)
            {
                crmIdValidationSelector.SelectedValue = CrmIdValidationPolicy.Policy.Strict;
            }

            this.PopulateDirectionsMenu(syncCallsMenu, Properties.Settings.Default.SyncCalls);
            this.PopulateDirectionsMenu(syncContactsMenu, Properties.Settings.Default.SyncContacts);
            this.PopulateDirectionsMenu(syncMeetingsMenu, Properties.Settings.Default.SyncMeetings);
            this.PopulateDirectionsMenu(syncTasksMenu, Properties.Settings.Default.SyncTasks);
        }

        /// <summary>
        /// Populate one of the synchronisation direction menus.
        /// </summary>
        /// <param name="directionMenu">The menu to populate.</param>
        /// <param name="setting">The value of the setting to set the menu value from.</param>
        private void PopulateDirectionsMenu(ComboBox directionMenu, SyncDirection.Direction setting)
        {
            var syncDirectionItems = Enum.GetValues(typeof(SyncDirection.Direction))
                    .Cast<SyncDirection.Direction>()
                    .Select(p => new { Key = (int)p, Value = SyncDirection.ToString(p) })
                    .OrderBy(o => o.Key)
                    .ToList();

            directionMenu.ValueMember = "Key";
            directionMenu.DisplayMember = "Value";
            directionMenu.DataSource = syncDirectionItems;

            try
            {
                directionMenu.SelectedValue = Convert.ToInt32(setting);
            }
            catch (NullReferenceException)
            {
                directionMenu.SelectedValue = SyncDirection.Direction.BiDirectional;
            }
        }

        private void GetAccountAutoArchivingSettings()
        {
            var settings = new EmailAccountsArchiveSettings();
            settings.Load();
            EmailArchiveAccountTabs.TabPages.Clear();
            var outlookSession = Application.Session;

            this.Log.Debug($"SettingsDialog: Setting up account archiving widget. Outlook version is {Globals.ThisAddIn.OutlookVersion}; there are {outlookSession.Accounts.Count} accounts");

            if (Globals.ThisAddIn.OutlookVersion >= OutlookMajorVersion.Outlook2010)
            {
                // Uses a Outlook 2013 APIs on Account object: DeliveryStore and GetRootFolder()
                // Needs work to make it work on Outlook 2010 and below.
                foreach (Account account in outlookSession.Accounts)
                {
                    this.Log.Debug($"SettingsDialog: Added email archiving tab for account {account.DisplayName}");
                    AddTabPage(account).LoadSettings(account, settings);
                }
            }
        }


        /// <summary>
        /// Create a tab containing a EmailAccountArchiveSettingsControl widget representing this outlook 
        /// account and add it to the tabs of the EmailAccountsArchiveSettings page.
        /// </summary>
        /// <param name="outlookAccount">The Outlook account to wrap.</param>
        /// <returns>The EmailAccountArchiveSettingsControl widget created</returns>
        private EmailAccountArchiveSettingsControl AddTabPage(Account outlookAccount)
        {
            var newPage = new TabPage();
            newPage.Text = outlookAccount.DisplayName;
            var pageControl = new EmailAccountArchiveSettingsControl();
            newPage.Controls.Add(pageControl);
            EmailArchiveAccountTabs.TabPages.Add(newPage);
            return pageControl;
        }

        private void SaveAccountAutoArchivingSettings()
        {
            var allSettings = EmailArchiveAccountTabs.TabPages.Cast<TabPage>()
                .SelectMany(tabPage => tabPage.Controls.OfType<EmailAccountArchiveSettingsControl>())
                .Select(accountSettingsControl => accountSettingsControl.SaveSettings())
                .ToList();

            var conbinedSettings = EmailAccountsArchiveSettings.Combine(allSettings);
            conbinedSettings.Save();
            Properties.Settings.Default.AutoArchive = conbinedSettings.HasAny;
        }

        private void frmSettings_FormClosing(object sender, FormClosingEventArgs e)
        {
            Globals.ThisAddIn.SuiteCRMUserSession.ClearAwaitingAuthentication();
        }

        private void btnTestLogin_Click(object sender, EventArgs e)
        {
            if (ValidateDetails())
            {
                try
                {
                    this.CheckUrlChanged();

                    using (WaitCursor.For(this))
                    {
                        Globals.ThisAddIn.Authenticate(SafelyGetText(txtURL), 
                            SafelyGetText(txtUsername), 
                            SafelyGetText(txtPassword), 
                            SafelyGetText(txtLDAPAuthenticationKey));
                    }
                    if (Globals.ThisAddIn.SuiteCRMUserSession.NotLoggedIn)
                    {
                        MessageBox.Show("Authentication failed!", "Authentication failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        MessageBox.Show("Login Successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    Properties.Settings.Default.Host = SafelyGetText(txtURL);
                    Properties.Settings.Default.Username = SafelyGetText(txtUsername);
                    Properties.Settings.Default.Password = SafelyGetText(txtPassword);
                }
                catch (Exception ex)
                {
                    Log.Error("Unable to connect to SuiteCRM", ex);
                    MessageBox.Show(ex.Message, "Unable to connect to SuiteCRM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void cbShowCustomModules_Click(object sender, EventArgs e)
        {
            if (cbShowCustomModules.Checked)
            {
                CustomModulesDialog objfrmCustomModules = new CustomModulesDialog();
                objfrmCustomModules.ShowDialog();
            }
        }

        private void chkEnableLDAPAuthentication_CheckedChanged(object sender, EventArgs e)
        {
            if (chkEnableLDAPAuthentication.Checked)
            {
                labelKey.Enabled = true;
                txtLDAPAuthenticationKey.Enabled = true;
                txtLDAPAuthenticationKey.Text = Properties.Settings.Default.LDAPKey;
            }
            else
            {
                labelKey.Enabled = false;
                txtLDAPAuthenticationKey.Enabled = false;
                txtLDAPAuthenticationKey.Text = string.Empty;
            }
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (cbShowCustomModules.Checked)
            {
                CustomModulesDialog objfrmCustomModules = new CustomModulesDialog();
                objfrmCustomModules.ShowDialog();
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!ValidateDetails())
            {
                this.DialogResult = DialogResult.None;
                return;
            }

            try
            {
                using (WaitCursor.For(this))
                {

                    ErrorHandler.DoOrHandleError(() => CheckUrlChanged(), "checking whether CRM URL has changed");

                    /* save settings before, and regardless of, test that login succeeds. 
                     * Otherwise in cases where login is impossible (e.g. network failure) 
                     * settings get lost. See bug #187 */
                    ErrorHandler.DoOrHandleError(() => this.SaveSettings(), "saving settings");

                    Globals.ThisAddIn.Authenticate();
                }

                if (Globals.ThisAddIn.SuiteCRMUserSession.NotLoggedIn)
                {
                    MessageBox.Show("Authentication failed!", "Authentication failed", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    this.DialogResult = DialogResult.None;
                    return;
                }
            }
            catch (Exception ex)
            {
                Log.Warn("Unable to connect to SuiteCRM", ex);
                MessageBox.Show(ex.Message, "Unable to connect to SuiteCRM", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                this.DialogResult = DialogResult.None;
                return;
            }

            RestAPIWrapper.FlushUserIdCache();

            base.Close();

            this.SettingsChanged?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Check whether the URL has changed; if it has, offer to clear down existing CRM ids.
        /// </summary>
        private void CheckUrlChanged()
        {
            var newUrl = SafelyGetText(txtURL);

            if (!newUrl.EndsWith(@"/"))
            {
                txtURL.Text = newUrl + "/";
                newUrl = SafelyGetText(txtURL);
            }
        }

        /// <summary>
        /// Return trimmed text from this box, but on no account throw an exception.
        /// </summary>
        /// <param name="box">The text box presumed to contain text.</param>
        /// <returns>The trimmed text</returns>
        private string SafelyGetText(TextBox box)
        {
            string result;

            try
            {
                result = box.Text == null ? String.Empty : box.Text.Trim();
            }
            catch (Exception)
            {
                result = string.Empty;
            }

            return result;
        }

        /// <summary>
        /// Save all settings from their current values in the dialog.
        /// </summary>
        private void SaveSettings()
        {
            if (!SafelyGetText(licenceText).Equals(Properties.Settings.Default.LicenceKey))
            {
                // if the license key has changed, force a new license check on next start.
                Properties.Settings.Default.LVSStartsRemaining = 0;
                Properties.Settings.Default.LVSLastStart = DateTime.MinValue;
            }


            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.Host = SafelyGetText(txtURL), "Saving Host");
            /* #6661: Forcing CurrentCrmIdPropertyName to "" will cause it to be recomputed from 
             * of the host URL */
            Properties.Settings.Default.CurrentCrmIdPropertyName = string.Empty;

            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.Username = SafelyGetText(txtUsername), "Saving Username");
            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.Password = SafelyGetText(txtPassword), "Saving Password");
            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.IsLDAPAuthentication = chkEnableLDAPAuthentication.Checked, "Saving IsLDAPAuthentication");
            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.LDAPKey = SafelyGetText(txtLDAPAuthenticationKey), "Saving LDAPKey");

            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.LicenceKey = SafelyGetText(licenceText), "Saving LicenceKey");

            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.ArchiveAttachments = this.cbEmailAttachments.Checked, "Saving ArchiveAttachments");
            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.AutomaticSearch = this.checkBoxAutomaticSearch.Checked, "Saving AutomaticSearch");
            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.ShowCustomModules = this.cbShowCustomModules.Checked, "Saving CRM URL");
            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.PopulateContextLookupList = this.checkBoxShowRightClick.Checked, "Saving PopulateContextLookupList");

            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.ExcludedEmails = this.SafelyGetText(txtAutoSync), "Saving ExcludedEmails");

            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.AutoArchiveFolders = new List<string>(), "Saving AutoArchiveFolders");

            ErrorHandler.DoOrHandleError(() => SaveAccountAutoArchivingSettings(), "Saving AccountAutoArchivingSettings");

            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.SyncCalls = (SyncDirection.Direction)this.syncCallsMenu.SelectedValue, "Saving SyncCalls");
            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.SyncMeetings = (SyncDirection.Direction)this.syncMeetingsMenu.SelectedValue, "Saving SyncMeetings");
            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.SyncTasks = (SyncDirection.Direction)this.syncTasksMenu.SelectedValue, "Saving SyncTasks");
            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.SyncContacts = (SyncDirection.Direction)this.syncContactsMenu.SelectedValue, "Saving SyncContacts");

            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.ShowConfirmationMessageArchive = this.chkShowConfirmationMessageArchive.Checked, "saving ShowConfirmationMessageArchive");
            ErrorHandler.DoOrHandleError(() => SaveSyncMaxRecords(), "saving SyncMaxRecords");

            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.LogLevel = (LogEntryType)logLevelSelector.SelectedValue, "Saving LogLevel");
            if (Globals.ThisAddIn.Log != null)
            {
                ErrorHandler.DoOrHandleError(
                    () => Log.Level = (LogEntryType) logLevelSelector.SelectedValue,
                    "Setting active log level");
            }

            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.ShowExceptions = (ErrorHandler.PopupWhen)showErrorsSelector.SelectedValue, "Saving ShowExceptions");

            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.StartupDeferral = (int)this.startupDeferralInput.Value, "Saving StartupDeferral");

            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.CrmIdValidationPolicy =
                (CrmIdValidationPolicy.Policy) crmIdValidationSelector.SelectedValue, "Saving CrmIdValidationPolicy");

            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.DaysOldEmailToAutoArchive =
                (int)Math.Ceiling(Math.Max((DateTime.Today - dtpAutoArchiveFrom.Value).TotalDays, 0)), "Saving DaysOldEmailToAutoArchive");

            ErrorHandler.DoOrHandleError(() => Properties.Settings.Default.Save(), "Storing all settings to disk");

            ErrorHandler.DoOrHandleError(() => Globals.ThisAddIn.StopUnconfiguredSynchronisationProcesses(), "stopping sync processes which are no longer required");
            ErrorHandler.DoOrHandleError(() => Globals.ThisAddIn.StartConfiguredSynchronisationProcesses(), "starting sync processes which are now required"); ;
        }

        private void SaveSyncMaxRecords()
        {
            Properties.Settings.Default.SyncMaxRecords = this.txtSyncMaxRecords.Text != string.Empty ?
                  Convert.ToInt32(this.txtSyncMaxRecords.Text):
                  0;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Reload();
            base.Close();
        }

        private void LinkToLogFileDir_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(ThisAddIn.LogDirPath);
        }

        private void advancedButton_Click(object sender, EventArgs e)
        {
            new AdvancedArchiveSettingsDialog().ShowDialog();
        }
    }
}
