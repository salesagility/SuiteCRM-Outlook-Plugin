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
    using Microsoft.Office.Interop.Outlook;
    using SuiteCRMClient.Logging;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Windows.Forms;
    using Exception = System.Exception;

    public partial class frmSettings : Form
    {
        public EventHandler SettingsChanged;

        public frmSettings()
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
            try
            {
                AddInTitleLabel.Text = ThisAddIn.AddInTitle;
                AddInVersionLabel.Text = "Version " + ThisAddIn.AddInVersion;

                if (Globals.ThisAddIn.SuiteCRMUserSession == null)
                {
                    Globals.ThisAddIn.SuiteCRMUserSession =
                        new SuiteCRMClient.UserSession(
                            "", "", "", "", Log, Globals.ThisAddIn.Settings.RestTimeout);
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

        private void LoadSettings()
        {
            if (Globals.ThisAddIn.Settings.host != "")
            {
                txtURL.Text = Globals.ThisAddIn.Settings.host;
                txtUsername.Text = Globals.ThisAddIn.Settings.username;
                txtPassword.Text = Globals.ThisAddIn.Settings.password;
                licenceText.Text = Globals.ThisAddIn.Settings.LicenceKey;
            }
            this.chkEnableLDAPAuthentication.Checked = Globals.ThisAddIn.Settings.IsLDAPAuthentication;
            this.txtLDAPAuthenticationKey.Text = Globals.ThisAddIn.Settings.LDAPKey;

            this.cbEmailAttachments.Checked = Globals.ThisAddIn.Settings.ArchiveAttachments;
            this.checkBoxAutomaticSearch.Checked = Globals.ThisAddIn.Settings.AutomaticSearch;
            this.cbShowCustomModules.Checked = Globals.ThisAddIn.Settings.ShowCustomModules;
            this.txtSyncMaxRecords.Text = Globals.ThisAddIn.Settings.SyncMaxRecords.ToString();
            this.checkBoxShowRightClick.Checked = Globals.ThisAddIn.Settings.PopulateContextLookupList;
            GetAccountAutoArchivingSettings();

            txtAutoSync.Text = Globals.ThisAddIn.Settings.ExcludedEmails;

            dtpAutoArchiveFrom.Value = DateTime.Now.AddDays(0 - Globals.ThisAddIn.Settings.DaysOldEmailToAutoArchive);
            chkShowConfirmationMessageArchive.Checked = Globals.ThisAddIn.Settings.ShowConfirmationMessageArchive;

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
            logLevelSelector.SelectedValue = Convert.ToInt32(Globals.ThisAddIn.Settings.LogLevel);

            var syncDirectionItems = Enum.GetValues(typeof(SyncDirection.Direction))
                    .Cast<SyncDirection.Direction>()
                    .Select(p => new { Key = (int)p, Value = SyncDirection.ToString(p) })
                    .OrderBy(o => o.Key)
                    .ToList();

            this.syncContactsMenu.ValueMember = "Key";
            this.syncContactsMenu.DisplayMember = "Value";
            this.syncCalendarMenu.ValueMember = "Key";
            this.syncCalendarMenu.DisplayMember = "Value";
            this.syncContactsMenu.DataSource = syncDirectionItems;
            this.syncCalendarMenu.DataSource = syncDirectionItems;
            this.syncCalendarMenu.SelectedValue = Convert.ToInt32(Globals.ThisAddIn.Settings.SyncCalendar);
            this.syncContactsMenu.SelectedValue = Convert.ToInt32(Globals.ThisAddIn.Settings.SyncContacts);
        }

        private void GetAccountAutoArchivingSettings()
        {
            var settings = new EmailAccountsArchiveSettings();
            settings.Load(Globals.ThisAddIn.Settings);
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
            conbinedSettings.Save(Globals.ThisAddIn.Settings);
            Globals.ThisAddIn.Settings.AutoArchive = conbinedSettings.HasAny;
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
                    if (txtLDAPAuthenticationKey.Text.Trim() == "")
                    {
                        txtLDAPAuthenticationKey.Text = null;
                    }
                    Globals.ThisAddIn.SuiteCRMUserSession = 
                        new SuiteCRMClient.UserSession(
                            txtURL.Text.Trim(), 
                            txtUsername.Text.Trim(), 
                            txtPassword.Text.Trim(), 
                            txtLDAPAuthenticationKey.Text.Trim(), 
                            Log, 
                            Globals.ThisAddIn.Settings.RestTimeout);

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
                    Globals.ThisAddIn.Settings.host = txtURL.Text.Trim();
                    Globals.ThisAddIn.Settings.username = txtUsername.Text.Trim();
                    Globals.ThisAddIn.Settings.password = txtPassword.Text.Trim();
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
                frmCustomModules objfrmCustomModules = new frmCustomModules();
                objfrmCustomModules.ShowDialog();
            }
        }

        private void chkEnableLDAPAuthentication_CheckedChanged(object sender, EventArgs e)
        {
            if (chkEnableLDAPAuthentication.Checked)
            {
                labelKey.Enabled = true;
                txtLDAPAuthenticationKey.Enabled = true;
                txtLDAPAuthenticationKey.Text = Globals.ThisAddIn.Settings.LDAPKey;
            }
            else
            {
                labelKey.Enabled = false;
                txtLDAPAuthenticationKey.Enabled = false;
                txtLDAPAuthenticationKey.Text = "";
            }
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (cbShowCustomModules.Checked)
            {
                frmCustomModules objfrmCustomModules = new frmCustomModules();
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
                if (txtLDAPAuthenticationKey.Text.Trim() == "")
                {
                    txtLDAPAuthenticationKey.Text = null;
                }
                Globals.ThisAddIn.SuiteCRMUserSession =
                    new SuiteCRMClient.UserSession(
                        txtURL.Text.Trim(),
                        txtUsername.Text.Trim(),
                        txtPassword.Text.Trim(),
                        txtLDAPAuthenticationKey.Text.Trim(),
                        Log,
                        Globals.ThisAddIn.Settings.RestTimeout);
                Globals.ThisAddIn.SuiteCRMUserSession.Login();
                if (Globals.ThisAddIn.SuiteCRMUserSession.NotLoggedIn)
                {
                    MessageBox.Show("Authentication failed!!!", "Authentication failed", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    this.DialogResult = DialogResult.None;
                    return;
                }
                Globals.ThisAddIn.Settings.host = txtURL.Text.Trim();
                Globals.ThisAddIn.Settings.username = txtUsername.Text.Trim();
                Globals.ThisAddIn.Settings.password = txtPassword.Text.Trim();
            }
            catch (Exception ex)
            {
                Log.Warn("Unable to connect to SuiteCRM", ex);
                MessageBox.Show(ex.Message, "Unable to connect to SuiteCRM", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                this.DialogResult = DialogResult.None;
                return;
            }

            Globals.ThisAddIn.Settings.IsLDAPAuthentication = chkEnableLDAPAuthentication.Checked;
            Globals.ThisAddIn.Settings.LDAPKey = txtLDAPAuthenticationKey.Text.Trim();

            Globals.ThisAddIn.Settings.LicenceKey = licenceText.Text.Trim();

            Globals.ThisAddIn.Settings.ArchiveAttachments = this.cbEmailAttachments.Checked;
            Globals.ThisAddIn.Settings.AutomaticSearch = this.checkBoxAutomaticSearch.Checked;
            Globals.ThisAddIn.Settings.ShowCustomModules = this.cbShowCustomModules.Checked;
            Globals.ThisAddIn.Settings.PopulateContextLookupList = this.checkBoxShowRightClick.Checked;

            Globals.ThisAddIn.Settings.ExcludedEmails = this.txtAutoSync.Text.Trim();

            Globals.ThisAddIn.Settings.AutoArchiveFolders = new List<string>();

            SaveAccountAutoArchivingSettings();

            Globals.ThisAddIn.Settings.SyncCalendar = (SyncDirection.Direction)this.syncCalendarMenu.SelectedValue;
            Globals.ThisAddIn.Settings.SyncContacts = (SyncDirection.Direction)this.syncContactsMenu.SelectedValue;

            Globals.ThisAddIn.Settings.ShowConfirmationMessageArchive = this.chkShowConfirmationMessageArchive.Checked;
            if (this.txtSyncMaxRecords.Text != string.Empty)
            {
                Globals.ThisAddIn.Settings.SyncMaxRecords = Convert.ToInt32(this.txtSyncMaxRecords.Text);
            }
            else
            {
                Globals.ThisAddIn.Settings.SyncMaxRecords = 0;
            }

            Globals.ThisAddIn.Settings.LogLevel = (LogEntryType)logLevelSelector.SelectedValue;
            Globals.ThisAddIn.Log.Level = Globals.ThisAddIn.Settings.LogLevel;

            Globals.ThisAddIn.Settings.DaysOldEmailToAutoArchive =
                (int)Math.Ceiling(Math.Max((DateTime.Today - dtpAutoArchiveFrom.Value).TotalDays, 0));

            Globals.ThisAddIn.Settings.Save();
            Globals.ThisAddIn.Settings.Reload();
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
