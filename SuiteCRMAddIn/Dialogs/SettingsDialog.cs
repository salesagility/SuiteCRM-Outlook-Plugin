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

        private bool ValidateDetails()
        {
            if (txtURL.Text.Trim().Length == 0)
            {
                MessageBox.Show("Please enter a valid SuiteCRM URL");
                txtURL.Focus();
                return false;
            }

            if (txtUsername.Text.Trim().Length == 0)
            {
                MessageBox.Show("Please enter a valid SuiteCRM Username");
                txtUsername.Focus();
                return false;
            }

            if (txtPassword.Text.Trim().Length == 0)
            {
                MessageBox.Show("Please enter a valid SuiteCRM Password");
                txtPassword.Focus();
                return false;
            }

            if (chkEnableLDAPAuthentication.Checked)
            {
                if (txtLDAPAuthenticationKey.Text.Trim().Length == 0)
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

                    if (Globals.ThisAddIn.SuiteCRMUserSession == null)
                    {
                        Globals.ThisAddIn.SuiteCRMUserSession =
                            new SuiteCRMClient.UserSession(
                                string.Empty, string.Empty, string.Empty, string.Empty, ThisAddIn.ProgId, Log, Properties.Settings.Default.RestTimeout);
                    }

                    Globals.ThisAddIn.SuiteCRMUserSession.AwaitingAuthentication = true;
                    LoadSettings();
                    LinkToLogFileDir.Text = ThisAddIn.LogDirPath;
                }
                catch (Exception ex)
                {
                    Log.Error("frmSettings_Load error", ex);
                    // Swallow exception!
                }
            }
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
            logLevelSelector.SelectedValue = Convert.ToInt32(Properties.Settings.Default.LogLevel);

            this.PopulateDirectionsMenu(syncCalendarMenu);
            this.PopulateDirectionsMenu(syncContactsMenu);

            this.syncCalendarMenu.SelectedValue = Convert.ToInt32(Properties.Settings.Default.SyncCalendar);
            this.syncContactsMenu.SelectedValue = Convert.ToInt32(Properties.Settings.Default.SyncContacts);
        }

        /// <summary>
        /// Populate one of the two synchronisation direction menus.
        /// </summary>
        /// <param name="directionMenu">The menu to populate.</param>
        private void PopulateDirectionsMenu(ComboBox directionMenu)
        {
            var syncDirectionItems = Enum.GetValues(typeof(SyncDirection.Direction))
                    .Cast<SyncDirection.Direction>()
                    .Select(p => new { Key = (int)p, Value = SyncDirection.ToString(p) })
                    .OrderBy(o => o.Key)
                    .ToList();

            directionMenu.ValueMember = "Key";
            directionMenu.DisplayMember = "Value";
            directionMenu.DataSource = syncDirectionItems;
        }

        private void GetAccountAutoArchivingSettings()
        {
            var settings = new EmailAccountsArchiveSettings();
            settings.Load();
            EmailArchiveAccountTabs.TabPages.Clear();
            var outlookSession = Application.Session;
            if (Globals.ThisAddIn.OutlookVersion >= OutlookMajorVersion.Outlook2013)
            {
                // Uses a Outlook 2013 APIs on Account object: DeliveryStore and GetRootFolder()
                // Needs work to make it work on Outlook 2010 and below.
                foreach (Account account in outlookSession.Accounts)
                {
                    var name = account.DisplayName;
                    var store = account.DeliveryStore;
                    var rootFolder = store.GetRootFolder();

                    var pageControl = AddTabPage(account);
                    pageControl.LoadSettings(account, settings);
                }
            }
        }

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
            Globals.ThisAddIn.SuiteCRMUserSession.AwaitingAuthentication = false;
        }

        private void btnTestLogin_Click(object sender, EventArgs e)
        {
            if (ValidateDetails())
            {
                try
                {
                    if (txtURL.Text.EndsWith(@"/"))
                    {
                    }
                    else
                    {
                        txtURL.Text = txtURL.Text + "/";
                    }
                    if (txtLDAPAuthenticationKey.Text.Trim() == string.Empty)
                    {
                        txtLDAPAuthenticationKey.Text = null;
                    }
                    Globals.ThisAddIn.SuiteCRMUserSession = 
                        new SuiteCRMClient.UserSession(
                            txtURL.Text.Trim(), 
                            txtUsername.Text.Trim(), 
                            txtPassword.Text.Trim(), 
                            txtLDAPAuthenticationKey.Text.Trim(), 
                            ThisAddIn.ProgId,
                            Log, 
                            Properties.Settings.Default.RestTimeout);

                    if (chkEnableLDAPAuthentication.Checked && txtLDAPAuthenticationKey.Text.Trim().Length != 0)
                    {
                        Globals.ThisAddIn.SuiteCRMUserSession.AuthenticateLDAP();
                    }
                    else
                    {
                        Globals.ThisAddIn.SuiteCRMUserSession.Login();
                    }
                    if (Globals.ThisAddIn.SuiteCRMUserSession.NotLoggedIn)
                    {
                        MessageBox.Show("Authentication failed!!!", "Authentication failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        MessageBox.Show("Login Successful!!!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    Properties.Settings.Default.Host = txtURL.Text.Trim();
                    Properties.Settings.Default.Username = txtUsername.Text.Trim();
                    Properties.Settings.Default.Password = txtPassword.Text.Trim();
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
                if (txtURL.Text.EndsWith(@"/"))
                {
                }
                else
                {
                    txtURL.Text = txtURL.Text + "/";
                }
                if (txtLDAPAuthenticationKey.Text.Trim() == string.Empty)
                {
                    txtLDAPAuthenticationKey.Text = null;
                }
                Globals.ThisAddIn.SuiteCRMUserSession =
                    new SuiteCRMClient.UserSession(
                        txtURL.Text.Trim(),
                        txtUsername.Text.Trim(),
                        txtPassword.Text.Trim(),
                        txtLDAPAuthenticationKey.Text.Trim(),
                        ThisAddIn.ProgId,
                        Log,
                        Properties.Settings.Default.RestTimeout);
                Globals.ThisAddIn.SuiteCRMUserSession.Login();
                if (Globals.ThisAddIn.SuiteCRMUserSession.NotLoggedIn)
                {
                    MessageBox.Show("Authentication failed!!!", "Authentication failed", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    this.DialogResult = DialogResult.None;
                    return;
                }
                Properties.Settings.Default.Host = txtURL.Text.Trim();
                Properties.Settings.Default.Username = txtUsername.Text.Trim();
                Properties.Settings.Default.Password = txtPassword.Text.Trim();
            }
            catch (Exception ex)
            {
                Log.Warn("Unable to connect to SuiteCRM", ex);
                MessageBox.Show(ex.Message, "Unable to connect to SuiteCRM", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                this.DialogResult = DialogResult.None;
                return;
            }

            Properties.Settings.Default.IsLDAPAuthentication = chkEnableLDAPAuthentication.Checked;
            Properties.Settings.Default.LDAPKey = txtLDAPAuthenticationKey.Text.Trim();

            Properties.Settings.Default.LicenceKey = licenceText.Text.Trim();

            Properties.Settings.Default.ArchiveAttachments = this.cbEmailAttachments.Checked;
            Properties.Settings.Default.AutomaticSearch = this.checkBoxAutomaticSearch.Checked;
            Properties.Settings.Default.ShowCustomModules = this.cbShowCustomModules.Checked;
            Properties.Settings.Default.PopulateContextLookupList = this.checkBoxShowRightClick.Checked;

            Properties.Settings.Default.ExcludedEmails = this.txtAutoSync.Text.Trim();

            Properties.Settings.Default.AutoArchiveFolders = new List<string>();

            SaveAccountAutoArchivingSettings();

            Properties.Settings.Default.SyncCalendar = (SyncDirection.Direction)this.syncCalendarMenu.SelectedValue;
            Properties.Settings.Default.SyncContacts = (SyncDirection.Direction)this.syncContactsMenu.SelectedValue;

            Properties.Settings.Default.ShowConfirmationMessageArchive = this.chkShowConfirmationMessageArchive.Checked;
            if (this.txtSyncMaxRecords.Text != string.Empty)
            {
                Properties.Settings.Default.SyncMaxRecords = Convert.ToInt32(this.txtSyncMaxRecords.Text);
            }
            else
            {
                Properties.Settings.Default.SyncMaxRecords = 0;
            }

            Properties.Settings.Default.LogLevel = (LogEntryType)logLevelSelector.SelectedValue;
            Globals.ThisAddIn.Log.Level = Properties.Settings.Default.LogLevel;

            Properties.Settings.Default.DaysOldEmailToAutoArchive =
                (int)Math.Ceiling(Math.Max((DateTime.Today - dtpAutoArchiveFrom.Value).TotalDays, 0));

            Properties.Settings.Default.Save();
            Properties.Settings.Default.Reload();

            clsSuiteCRMHelper.FlushUserIdCache();

            base.Close();

            this.SettingsChanged?.Invoke(this, EventArgs.Empty);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void LinkToLogFileDir_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(ThisAddIn.LogDirPath);
        }
    }
}
