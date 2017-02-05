using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn
{
    using BusinessLogic;
    using SuiteCRMClient.Logging;

    public partial class EmailAccountArchiveSettingsControl : UserControl
    {
        private readonly Outlook.Account _outlookAccount;

        public EmailAccountArchiveSettingsControl(Outlook.Account outlookAccount)
        {
            _outlookAccount = outlookAccount;
            InitializeComponent();
        }

        protected ILogger Log => Globals.ThisAddIn.Log;

        public void LoadSettings(EmailAccountsArchiveSettings settings)
        {
            var store = _outlookAccount.DeliveryStore;
            var rootFolder = store.GetRootFolder();
            var smtpAddress = _outlookAccount.SmtpAddress;
            ArchiveInboundCheckbox.Text = $"Archive Mail Received by {smtpAddress}";
            ArchiveOutboundCheckbox.Text = $"Archive Mail Sent by {smtpAddress}";

            var storeId = store.StoreID;
            GetMailFolders(settings, rootFolder.Folders);
            ArchiveInboundCheckbox.Checked = settings.AccountsToArchiveInbound.Contains(storeId);
            ArchiveOutboundCheckbox.Checked = settings.AccountsToArchiveOutbound.Contains(storeId);
        }

        public EmailAccountsArchiveSettings SaveSettings()
        {
            var store = _outlookAccount.DeliveryStore;
            var result = new EmailAccountsArchiveSettings().Clear();

            var storeId = store.StoreID;
            GetCheckedFoldersHelper(tsResults.Nodes, result.SelectedFolderEntryIds);
            if (ArchiveInboundCheckbox.Checked)
                result.AccountsToArchiveInbound.Add(storeId);
            if (ArchiveOutboundCheckbox.Checked)
                result.AccountsToArchiveOutbound.Add(storeId);

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
                Log.Error("GetCheckedFolders error", ex);
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


    }
}
