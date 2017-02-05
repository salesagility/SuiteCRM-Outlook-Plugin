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
using System.Diagnostics;
using System.Windows.Forms;
using SuiteCRMClient;
using SuiteCRMClient.Logging;

namespace SuiteCRMAddIn
{
    using Microsoft.Office.Interop.Outlook;
    using Exception = System.Exception;

    public partial class frmSettings : Form
    {
        private clsSettings settings = Globals.ThisAddIn.settings;

        public EventHandler SettingsChanged;

        public frmSettings()
        {
            InitializeComponent();
        }

        private ILogger Log => Globals.ThisAddIn.Log;

        private void GetCheckedFolders(TreeNode objInpNode)
        {
            try
            {
                if (objInpNode.Checked)
                    this.settings.AutoArchiveFolders.Add(objInpNode.Tag.ToString());

                foreach (TreeNode objNode in objInpNode.Nodes)
                {
                    if (objNode.Nodes.Count > 0)
                    {
                        GetCheckedFolders(objNode);
                    }
                    else
                    {
                        if (objNode.Checked)
                            this.settings.AutoArchiveFolders.Add(objNode.Tag.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                // Suppress exception.
                Log.Error("GetCheckedFolders error", ex);
            }
        }

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
                AddInVersionLabel.Text = "version " + ThisAddIn.AddInVersion;

                if (Globals.ThisAddIn.SuiteCRMUserSession == null)
                    Globals.ThisAddIn.SuiteCRMUserSession = new SuiteCRMClient.clsUsersession("", "", "", "", Log);

                Globals.ThisAddIn.SuiteCRMUserSession.AwaitingAuthentication = true;
                LoadSettings();
                LinkToLogFileDir.Text = ThisAddIn.LogDirPath;
                UpdateUIState();
            }
            catch (Exception ex)
            {
                Log.Error("frmSettings_Load error", ex);
                // Swallow exception!
            }
        }

        private void LoadSettings()
        {
            if (settings.host != "")
            {
                txtURL.Text = settings.host;
                txtUsername.Text = settings.username;
                txtPassword.Text = settings.password;
            }
            this.chkEnableLDAPAuthentication.Checked = this.settings.IsLDAPAuthentication;
            this.txtLDAPAuthenticationKey.Text = this.settings.LDAPKey;

            this.cbEmailAttachments.Checked = settings.ArchiveAttachmentsDefault;
            this.checkBoxAutomaticSearch.Checked = true;
            this.cbShowCustomModules.Checked = settings.ShowCustomModules;
            this.txtSyncMaxRecords.Text = this.settings.SyncMaxRecords.ToString();
            this.checkBoxShowRightClick.Checked = this.settings.PopulateContextLookupList;
            this.chkAutoArchive.Checked = this.settings.AutoArchive;
            this.chkSyncCalendar.Checked = this.settings.SyncCalendar;
            this.chkSyncContacts.Checked = this.settings.SyncContacts;
            this.tsResults.AfterCheck += new TreeViewEventHandler(this.tree_search_results_AfterCheck);
            this.tsResults.AfterExpand += new TreeViewEventHandler(this.tree_search_results_AfterExpand);
            this.tsResults.NodeMouseClick += new TreeNodeMouseClickEventHandler(this.tree_search_results_NodeMouseClick);
            this.tsResults.Nodes.Clear();
            this.tsResults.CheckBoxes = true;
            GetMailFolders(Globals.ThisAddIn.Application.Session.Folders);
            this.tsResults.ExpandAll();

            txtAutoSync.Text = settings.ExcludedEmails;

            gbFirstTime.Visible = settings.IsFirstTime;
            dtpAutoArchiveFrom.Value = System.DateTime.Now.AddDays(-10);
            chkShowConfirmationMessageArchive.Checked = this.settings.ShowConfirmationMessageArchive;

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

            DetailedLoggingCheckBox.Checked = settings.LogLevel <= LogEntryType.Debug;
        }

        private void GetMailFolders(Folders folders)
        {
            GetMailFolders(folders, tsResults.Nodes);
        }

        private void GetMailFolders(Folders folders, TreeNodeCollection nodes)
        {
            try
            {
                foreach (Folder objFolder in folders)
                {
                    var objNode = new TreeNode() { Tag = objFolder.EntryID, Text = objFolder.Name };
                    if (this.settings.AutoArchiveFolders.Contains(objFolder.EntryID))
                        objNode.Checked = true;
                    nodes.Add(objNode);
                    var nestedFolders = objFolder.Folders;
                    if (nestedFolders.Count > 0)
                    {
                        GetMailFolders(nestedFolders, objNode.Nodes);
                    }
                }
            }
            catch (Exception ex)
            {
                // Swallow exception(!)
                Log.Error("GetMailFolders error", ex);
            }
        }

        private void tree_search_results_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if ((e.Action != TreeViewAction.Unknown) && (e.Node.Nodes.Count > 0))
            {
                this.CheckAllChildNodes(e.Node, e.Node.Checked);
            }
        }
        private void CheckAllChildNodes(TreeNode treeNode, bool nodeChecked)
        {
            foreach (TreeNode node in treeNode.Nodes)
            {
                node.Checked = nodeChecked;
                if (node.Nodes.Count > 0)
                {
                    this.CheckAllChildNodes(node, nodeChecked);
                }
            }
        }

        private void tree_search_results_AfterExpand(object sender, TreeViewEventArgs e)
        {
            TreeViewAction action = e.Action;
        }

        private void tree_search_results_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (((e.Button == MouseButtons.Right) && (e.Node.Tag.ToString() != "root_node")) && (e.Node.Tag.ToString() != "sub_root_node"))
            {
                this.tsResults.SelectedNode = e.Node;
            }
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
                    Globals.ThisAddIn.SuiteCRMUserSession = new SuiteCRMClient.clsUsersession(txtURL.Text.Trim(), txtUsername.Text.Trim(), txtPassword.Text.Trim(), txtLDAPAuthenticationKey.Text.Trim(), Log);

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
                    settings.host = txtURL.Text.Trim();
                    settings.username = txtUsername.Text.Trim();
                    settings.password = txtPassword.Text.Trim();
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
                txtLDAPAuthenticationKey.Text = settings.LDAPKey;
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
                Globals.ThisAddIn.SuiteCRMUserSession = new SuiteCRMClient.clsUsersession(txtURL.Text.Trim(),
                    txtUsername.Text.Trim(), txtPassword.Text.Trim(), txtLDAPAuthenticationKey.Text.Trim(), Log);
                Globals.ThisAddIn.SuiteCRMUserSession.Login();
                if (Globals.ThisAddIn.SuiteCRMUserSession.NotLoggedIn)
                {
                    MessageBox.Show("Authentication failed!!!", "Authentication failed", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    this.DialogResult = DialogResult.None;
                    return;
                }
                settings.host = txtURL.Text.Trim();
                settings.username = txtUsername.Text.Trim();
                settings.password = txtPassword.Text.Trim();
            }
            catch (Exception ex)
            {
                Log.Warn("Unable to connect to SuiteCRM", ex);
                MessageBox.Show(ex.Message, "Unable to connect to SuiteCRM", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                this.DialogResult = DialogResult.None;
                return;
            }

            settings.IsLDAPAuthentication = chkEnableLDAPAuthentication.Checked;
            settings.LDAPKey = txtLDAPAuthenticationKey.Text.Trim();

            settings.ArchiveAttachmentsDefault = this.cbEmailAttachments.Checked;
            settings.AutomaticSearch = true;
            settings.ShowCustomModules = this.cbShowCustomModules.Checked;
            settings.PopulateContextLookupList = this.checkBoxShowRightClick.Checked;

            settings.ExcludedEmails = this.txtAutoSync.Text.Trim();

            settings.AutoArchiveFolders = new List<string>();

            foreach (TreeNode objNode in this.tsResults.Nodes)
            {
                if (objNode.Nodes.Count > 0)
                {
                    GetCheckedFolders(objNode);
                }
            }

            settings.AutoArchive = this.chkAutoArchive.Checked;
            settings.SyncCalendar = this.chkSyncCalendar.Checked;
            settings.SyncContacts = this.chkSyncContacts.Checked;
            settings.ShowConfirmationMessageArchive = this.chkShowConfirmationMessageArchive.Checked;
            if (this.txtSyncMaxRecords.Text != string.Empty)
            {
                this.settings.SyncMaxRecords = Convert.ToInt32(this.txtSyncMaxRecords.Text);
            }
            else
            {
                this.settings.SyncMaxRecords = 0;
            }

            settings.LogLevel = DetailedLoggingCheckBox.Checked ? LogEntryType.Debug : LogEntryType.Information;

            if (settings.IsFirstTime)
            {
                settings.IsFirstTime = false;
                System.Threading.Thread objThread =
                    new System.Threading.Thread(() => Globals.ThisAddIn.ProcessMails(dtpAutoArchiveFrom.Value));
                objThread.Start();
            }

            this.settings.Save();
            this.settings.Reload();
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

        private void chkAutoArchive_CheckedChanged(object sender, EventArgs e)
        {
            UpdateUIState();
        }

        private void UpdateUIState()
        {
            LimitArchivingLabel.Enabled = chkAutoArchive.Checked;
            tsResults.Enabled = chkAutoArchive.Checked;
        }
    }
}
