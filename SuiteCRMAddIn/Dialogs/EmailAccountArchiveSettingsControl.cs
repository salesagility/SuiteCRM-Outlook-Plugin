using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.Dialogs
{
    using BusinessLogic;
    using SuiteCRMClient.Logging;

    public partial class EmailAccountArchiveSettingsControl : UserControl
    {
        private string _outlookStoreId;

        public EmailAccountArchiveSettingsControl()
        {
            InitializeComponent();
        }

        protected ILogger Log => Globals.ThisAddIn.Log;

        public void LoadSettings(Outlook.Account outlookAccount, EmailAccountsArchiveSettings settings)
        {
            var store = outlookAccount.DeliveryStore;
            _outlookStoreId = store.StoreID;
            var rootFolder = store.GetRootFolder();
            var smtpAddress = outlookAccount.SmtpAddress;
            ArchiveInboundCheckbox.Text = $"Archive Mail Received by {smtpAddress}";
            ArchiveOutboundCheckbox.Text = $"Archive Mail Sent by {smtpAddress}";

            GetMailFolders(settings, rootFolder.Folders);
            ArchiveInboundCheckbox.Checked = settings.AccountsToArchiveInbound.Contains(_outlookStoreId);
            ArchiveOutboundCheckbox.Checked = settings.AccountsToArchiveOutbound.Contains(_outlookStoreId);
        }

        public EmailAccountsArchiveSettings SaveSettings()
        {
            var result = new EmailAccountsArchiveSettings();
            if (_outlookStoreId != null)
            {
                GetCheckedFoldersHelper(tsResults.Nodes, result.SelectedFolderEntryIds);
                if (ArchiveInboundCheckbox.Checked)
                    result.AccountsToArchiveInbound.Add(_outlookStoreId);
                if (ArchiveOutboundCheckbox.Checked)
                    result.AccountsToArchiveOutbound.Add(_outlookStoreId);
            }

            return result;
        }

        private ISet<string> GetCheckedFolders(TreeNodeCollection nodes)
        {
            var results = new HashSet<string>();
            GetCheckedFoldersHelper(nodes, results);
            return results;
        }

        private void GetCheckedFoldersHelper(TreeNodeCollection nodes, HashSet<string> results)
        {
            try
            {
                foreach (TreeNode node in nodes)
                {
                    if (node.Checked)
                    {
                        results.Add(node.Tag.ToString());
                    }
                    var childNodes = node.Nodes;
                    if (childNodes.Count > 0)
                    {
                        GetCheckedFoldersHelper(childNodes, results);
                    }
                }
            }
            catch (Exception ex)
            {
                // Suppress exception.
                ErrorHandler.Handle("Failed while fetching checked folders", ex);
            }
        }

        private void GetMailFolders(EmailAccountsArchiveSettings settings, Outlook.Folders folders)
        {
            this.tsResults.Nodes.Clear();
            this.tsResults.CheckBoxes = true;
            GetMailFolders(folders, tsResults.Nodes, settings.SelectedFolderEntryIds);
            this.tsResults.ExpandAll();
        }

        private void GetMailFolders(Outlook.Folders folders, TreeNodeCollection nodes, ISet<string> selectedFolderEntryIds)
        {
            try
            {
                foreach (Outlook.Folder objFolder in folders)
                {
                    var objNode = new TreeNode() { Tag = objFolder.EntryID, Text = objFolder.Name };
                    if (selectedFolderEntryIds.Contains(objFolder.EntryID))
                        objNode.Checked = true;
                    nodes.Add(objNode);
                    var nestedFolders = objFolder.Folders;
                    if (nestedFolders.Count > 0)
                    {
                        GetMailFolders(nestedFolders, objNode.Nodes, selectedFolderEntryIds);
                    }
                }
            }
            catch (Exception ex)
            {
                // Swallow exception(!)
                ErrorHandler.Handle("Failed while getting email folders", ex);
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


    }
}
