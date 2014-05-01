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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.Office.Interop;

namespace SuiteCRMOutlookAddIn
{
    public partial class frmSettings : Form
    {
        private clsSettings settings = AddinModule.CurrentInstance.settings;
        public frmSettings()
        {
            InitializeComponent();
        }
        
        private void GetCheckedFolders(TreeNode objInpNode)
        {
            if (objInpNode.Checked)
                this.settings.auto_archive_folders.Add(objInpNode.Tag.ToString());

            foreach (TreeNode objNode in objInpNode.Nodes)
            {
                if (objNode.Nodes.Count > 0)
                {
                    GetCheckedFolders(objNode);
                }
                else
                {
                    if (objNode.Checked)
                        this.settings.auto_archive_folders.Add(objNode.Tag.ToString());
                }
            }
        }

        private bool ValidateDetails()
        {
            if (txtURL.Text.Trim().Length == 0)
            {
                MessageBox.Show("Please enter a valid SugarCRM URL");
                txtURL.Focus();
                return false;
            }

            if (txtUsername.Text.Trim().Length == 0)
            {
                MessageBox.Show("Please enter a valid SugarCRM Username");
                txtUsername.Focus();
                return false;
            }

            if (txtPassword.Text.Trim().Length == 0)
            {
                MessageBox.Show("Please enter a valid SugarCRM Password");
                txtPassword.Focus();
                return false;
            }

            return true;
        }

        private void frmSettings_Load(object sender, EventArgs e)
        {
            if (AddinModule.CurrentInstance.SugarCRMUserSession == null)
                AddinModule.CurrentInstance.SugarCRMUserSession = new SuiteCRMClient.clsUsersession("", "", "");

            AddinModule.CurrentInstance.SugarCRMUserSession.AwaitingAuthentication = true;
            if (settings.host != "")
            {
                txtURL.Text = settings.host;
                txtUsername.Text = settings.username;
                txtPassword.Text = settings.password;
            }
            this.cbEmailAttachments.Checked =settings.archive_attachments_default;
            this.checkBoxAutomaticSearch.Checked = true;
            this.cbShowCustomModules.Checked = settings.show_custom_modules;
            this.txtSyncMaxRecords.Text = this.settings.sync_max_records.ToString();
            this.checkBoxShowRightClick.Checked = this.settings.populate_context_lookup_list;
            this.chkAutoArchive.Checked = this.settings.auto_archive;
            this.tsResults.AfterCheck += new TreeViewEventHandler(this.tree_search_results_AfterCheck);
            this.tsResults.AfterExpand += new TreeViewEventHandler(this.tree_search_results_AfterExpand);
            this.tsResults.NodeMouseClick += new TreeNodeMouseClickEventHandler(this.tree_search_results_NodeMouseClick);
            this.tsResults.Nodes.Clear();
            this.tsResults.CheckBoxes = true;
            foreach (Microsoft.Office.Interop.Outlook.Folder objFolder in AddinModule.CurrentInstance.OutlookApp.Session.Folders)
            {
                if (objFolder.Name.ToUpper() == "SENT ITEMS" || objFolder.Name.ToUpper() == "OUTBOX")
                    continue;

                TreeNode objNode = new TreeNode() { Tag = objFolder.EntryID, Text = objFolder.Name };                
                if (this.settings.auto_archive_folders.Contains(objFolder.EntryID))
                    objNode.Checked = true;
                tsResults.Nodes.Add(objNode);
                if (objFolder.Folders.Count > 0)
                {
                    GetMailFolders(objFolder, objNode);
                }
            }
            this.tsResults.ExpandAll();

            txtAuotSync.Text = settings.ExcludedEmails;

            gbFirstTime.Visible = settings.IsFirstTime;
            dtpAutoArchiveFrom.Value = System.DateTime.Now.AddDays(-10);

            settings.AttachmentsChecked = cbEmailAttachments.Checked;
        }

        private void GetMailFolders(Microsoft.Office.Interop.Outlook.Folder objInpFolder, TreeNode objInpNode)
        {
            foreach (Microsoft.Office.Interop.Outlook.Folder objFolder in objInpFolder.Folders)
            {
                if (objFolder.Name.ToUpper() == "SENT ITEMS" || objFolder.Name.ToUpper() == "OUTBOX")
                    continue;
                if (objFolder.Folders.Count > 0)
                {
                    TreeNode objNode = new TreeNode() { Tag = objFolder.EntryID, Text = objFolder.Name };                    
                    if (this.settings.auto_archive_folders.Contains(objFolder.EntryID))
                        objNode.Checked = true;
                    objInpNode.Nodes.Add(objNode);
                    GetMailFolders(objFolder, objNode);
                }
                else
                {
                    TreeNode objNode = new TreeNode() { Tag = objFolder.EntryID, Text = objFolder.Name };
                    if (this.settings.auto_archive_folders.Contains(objFolder.EntryID))
                        objNode.Checked = true;
                    objInpNode.Nodes.Add(objNode);
                }
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
            AddinModule.CurrentInstance.SugarCRMUserSession.AwaitingAuthentication = false;
        }

        private void btnTestLogin_Click(object sender, EventArgs e)
        {
            if (ValidateDetails())
            {
                try
                {
                    AddinModule.CurrentInstance.SugarCRMUserSession = new SuiteCRMClient.clsUsersession(txtURL.Text.Trim(), txtUsername.Text.Trim(), txtPassword.Text.Trim());
                    AddinModule.CurrentInstance.SugarCRMUserSession.Login();
                    if (AddinModule.CurrentInstance.SugarCRMUserSession.id == "")
                    {
                        MessageBox.Show("Authentication failed!!!", "Authentication failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        MessageBox.Show("Login Succeed!!!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    settings.host = txtURL.Text.Trim();
                    settings.username = txtUsername.Text.Trim();
                    settings.password = txtPassword.Text.Trim();                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Unable to connect to SugarCRM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ex.Data.Clear();                    
                }
            }
        }

        
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (ValidateDetails())
            {
                try
                {
                    AddinModule.CurrentInstance.SugarCRMUserSession = new SuiteCRMClient.clsUsersession(txtURL.Text.Trim(), txtUsername.Text.Trim(), txtPassword.Text.Trim());
                    AddinModule.CurrentInstance.SugarCRMUserSession.Login();
                    if (AddinModule.CurrentInstance.SugarCRMUserSession.id == "")
                    {
                        MessageBox.Show("Authentication failed!!!", "Authentication failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.DialogResult = DialogResult.None;
                        return;
                    }
                    settings.host = txtURL.Text.Trim();
                    settings.username = txtUsername.Text.Trim();
                    settings.password = txtPassword.Text.Trim();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Unable to connect to SugarCRM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ex.Data.Clear();
                    this.DialogResult = DialogResult.None;
                    return;
                }
                settings.archive_attachments_default = this.cbEmailAttachments.Checked;
                settings.automatic_search = true;
                settings.show_custom_modules = this.cbShowCustomModules.Checked;
                settings.populate_context_lookup_list = this.checkBoxShowRightClick.Checked;

                settings.ExcludedEmails = this.txtAuotSync.Text.Trim();

                settings.auto_archive_folders = new List<string>();

                foreach (TreeNode objNode in this.tsResults.Nodes)
                {
                    if (objNode.Nodes.Count > 0)
                    {
                        GetCheckedFolders(objNode);
                    }
                }

                if (settings.auto_archive != this.chkAutoArchive.Checked && settings.IsFirstTime == false)
                {
                    System.Threading.Thread objThread = new System.Threading.Thread(() => AddinModule.CurrentInstance.ProcessMails());
                    objThread.Start();
                }

                settings.auto_archive = this.chkAutoArchive.Checked;

                if (this.txtSyncMaxRecords.Text != string.Empty)
                {
                    this.settings.sync_max_records = Convert.ToInt32(this.txtSyncMaxRecords.Text);
                }
                else
                {
                    this.settings.sync_max_records = 0;
                }
                if (settings.IsFirstTime)
                {
                    settings.IsFirstTime = false;
                    System.Threading.Thread objThread = new System.Threading.Thread(() => AddinModule.CurrentInstance.ProcessMails(dtpAutoArchiveFrom.Value));
                    objThread.Start();
                }
                this.settings.Save();
                this.settings.Reload();
                base.Close();
            }
            else
                this.DialogResult = DialogResult.None;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            base.Close();
        }

      
        private void cbShowCustomModules_Click(object sender, EventArgs e)
        {
            if (cbShowCustomModules.Checked)
            {
                frmCustomModules objfrmCustomModules = new frmCustomModules();
                objfrmCustomModules.ShowDialog();
            }
        }
    }
}
