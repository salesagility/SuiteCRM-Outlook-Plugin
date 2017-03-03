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
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SuiteCRMClient.RESTObjects;
using SuiteCRMClient;
using System.Collections.Specialized;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Web;

namespace SuiteCRMAddIn
{
    using BusinessLogic;
    using SuiteCRMClient.Email;
    using SuiteCRMClient.Exceptions;
    using SuiteCRMClient.Logging;
    using Exception = System.Exception;

    public partial class frmArchive : Form
    {

        public frmArchive()
        {
            InitializeComponent();
        }

        private clsSettings settings = Globals.ThisAddIn.Settings;
        public string type;

        private ILogger Log => Globals.ThisAddIn.Log;

        private void GetCustomModules()
        {
            if (this.settings.CustomModules != null)
            {
                StringEnumerator enumerator = this.settings.CustomModules.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    string[] strArray = enumerator.Current.Split(new char[] { '|' });
                    ListViewItem item = new ListViewItem
                    {
                        Tag = strArray[0],
                        Text = strArray[1],
                    };
                    if (strArray[0] != "None" || strArray[1] != "None")
                        this.lstViewSearchModules.Items.Add(item);
                }
            }
        }


        private void frmArchive_Load(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Settings.ShowCustomModules)
            {
                this.GetCustomModules();
            }
            try
            {
                foreach (string str in this.settings.SelectedSearchModules.Split(new char[] { ',' }))
                {
                    int num = Convert.ToInt32(str);
                    this.lstViewSearchModules.Items[num].Checked = true;
                }
            }
            catch (System.Exception)
            {
                // Swallow exception(!)
            }

            this.tsResults.AfterCheck += new TreeViewEventHandler(this.tsResults_AfterCheck);
            this.tsResults.AfterExpand += new TreeViewEventHandler(this.tsResults_AfterExpand);
            this.tsResults.NodeMouseClick += new TreeNodeMouseClickEventHandler(this.tsResults_NodeMouseClick);
            this.txtSearch.KeyDown += new KeyEventHandler(this.txtSearch_KeyDown);
            this.lstViewSearchModules.ItemChecked += new ItemCheckedEventHandler(this.lstViewSearchModules_ItemChecked);
            base.FormClosed += new FormClosedEventHandler(this.frmArchive_FormClosed);

            this.txtSearch.Text = ConstructSearchText(Globals.ThisAddIn.SelectedEmails);

            if (this.settings.AutomaticSearch)
            {
                this.btnSearch_Click(null, null);
            }
        }

        /// <summary>
        /// Construct suitable search text from this list of emails.
        /// </summary>
        /// <param name="emails">A list of emails, presumably those selected by the user</param>
        /// <returns>A string comprising the sender addresses from the emails, comma separated.</returns>
        private static string ConstructSearchText(IEnumerable<MailItem> emails)
        {
            List<string> addresses = new List<string>();
            string searchText = String.Empty;

            foreach (var email in emails)
            {
                addresses.Add(clsGlobals.GetSMTPEmailAddress(email));
            }

            foreach (var address in addresses.OrderBy(x => x).GroupBy(x => x).Select(g => g.First()))
            {
                searchText += address + ",";
            }

            return searchText.TrimEnd(',');
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

        public void btnSearch_Click(object sender, EventArgs e)
        {
            this.tsResults.Nodes.Clear();

            if (this.txtSearch.Text.Contains<char>(','))
            {
                foreach (string str in this.txtSearch.Text.Split(new char[] { ',' }).OrderBy(x => x).GroupBy(x => x).Select(g => g.First()))
                {
                    this.Search(str);
                }
            }
            if (this.txtSearch.Text.Contains(";"))
            {
                foreach (string str2 in this.txtSearch.Text.Split(new char[] { ';' }))
                {
                    this.Search(str2);
                }
            }
            else
            {
                this.Search(this.txtSearch.Text);
            }
        }

        private bool UnallowedNumber(string strText)
        {
            char[] charUnallowedNumber = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
            for (int i = 0; i < charUnallowedNumber.Length; i++)
                if (strText.StartsWith(charUnallowedNumber[i].ToString()))
                    return true;
            return false;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        public void Search(string searchText)
        {
            using (WaitCursor.For(this))
                try
                {
                    List<string> list = new List<string> { "Accounts", "Contacts", "Leads", "Bugs", "Projects", "Cases", "Opportunties" };
                    this.tsResults.CheckBoxes = true;
                    if (searchText == string.Empty)
                    {
                        MessageBox.Show("Please enter some text to search", "Invalid search", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        searchText = searchText.TrimStart(new char[0]);
                        string[] strArray = searchText.Split(new char[] { ' ' });
                        string usString = strArray[0];
                        string str2 = string.Empty;
                        string str3 = "OR";
                        if (strArray.Length > 1)
                        {
                            str2 = strArray[1];
                            str3 = "AND";
                        }
                        else
                        {
                            str2 = strArray[0];
                        }
                        foreach (ListViewItem item in this.lstViewSearchModules.Items)
                        {
                            try
                            {
                                TreeNode node;
                                eGetEntryListResult queryResult;
                                if (!item.Checked)
                                {
                                    continue;
                                }
                                string moduleName = item.Tag.ToString();

                                if (moduleName != "All")
                                {
                                    if (this.tsResults.Nodes[moduleName] == null)
                                    {
                                        node = new TreeNode(moduleName)
                                        {
                                            Tag = "root_node",
                                            Name = moduleName
                                        };
                                        this.tsResults.Nodes.Add(node);
                                    }
                                    else
                                    {
                                        node = this.tsResults.Nodes[moduleName];
                                    }
                                    string[] fields = new string[6];
                                    fields[0] = "id";
                                    fields[1] = "first_name";
                                    fields[2] = "last_name";
                                    fields[3] = "name";

                                    string queryText = ConstructQueryTextForModuleName(searchText, usString, str2, str3, moduleName, fields);

                                    try
                                    {
                                        queryResult = clsSuiteCRMHelper.GetEntryList(moduleName, queryText, settings.SyncMaxRecords, "date_entered DESC", 0, false, fields);
                                    }
                                    catch (System.Exception any)
                                    {
                                        Globals.ThisAddIn.Log.Error($"Failure when custom module included (1)\n\tQuery was '{queryText}'", any);
                                        // Swallow exception(!)
                                        try {
                                            queryResult = clsSuiteCRMHelper.GetEntryList(moduleName, queryText.Replace("%", ""), settings.SyncMaxRecords, "date_entered DESC", 0, false, fields);
                                        }
                                        catch (Exception secondFail)
                                        {
                                            queryText = queryText.Replace("%", "");
                                            Globals.ThisAddIn.Log.Error($"Failure when custom module included (2)\n\tQuery was '{queryText}'", secondFail);
                                            MessageBox.Show(
                                                $"An error was encountered while querying module '{moduleName}'. The error has been logged",
                                                "Query error",
                                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                            queryResult = null;
                                        }
                                    }
                                    if (queryResult != null && queryResult.result_count > 0)
                                    {
                                        this.populateTree(queryResult, moduleName, node);
                                    }
                                    else if (!list.Contains(moduleName) && clsSuiteCRMHelper.GetFields(moduleName).Contains("first_name"))
                                    {
                                        queryText = "(" + moduleName.ToLower() + ".first_name LIKE '%" + clsGlobals.MySqlEscape(usString) + "%' " + str3 + " " + moduleName.ToLower() + ".last_name LIKE '%" + clsGlobals.MySqlEscape(str2) + "%')  OR (" + moduleName.ToLower() + ".id in (select eabr.bean_id from email_addr_bean_rel eabr INNER JOIN email_addresses ea on eabr.email_address_id = ea.id where eabr.bean_module = '" + moduleName + "' and ea.email_address LIKE '%" + clsGlobals.MySqlEscape(searchText) + "%'))";
                                        eGetEntryListResult _result2 = clsSuiteCRMHelper.GetEntryList(moduleName, queryText, settings.SyncMaxRecords, "date_entered DESC", 0, false, fields);
                                        if (_result2.result_count > 0)
                                        {
                                            this.populateTree(_result2, moduleName, node);
                                        }
                                    }
                                    if (node.GetNodeCount(true) <= 0)
                                    {
                                        node.Remove();
                                    }
                                }
                            }
                            catch (System.Exception any)
                            {
                                Globals.ThisAddIn.Log.Error("Failure when custom module included (3)", any);

                                // Swallow exception(!)
                                this.tsResults.Nodes.Clear();
                            }
                        }
                        if (this.tsResults.Nodes.Count <= 0)
                        {
                            TreeNode node2 = new TreeNode("No results found")
                            {
                                Name = "No results",
                                Text = "No Result"
                            };
                            this.tsResults.Nodes.Add(node2);
                            this.tsResults.CheckBoxes = false;
                        }
                        this.txtSearch.Enabled = true;
                    }
                }
                catch (System.Exception)
                {
                    // Swallow exception(!)

                    this.tsResults.Nodes.Clear();
                    TreeNode node2 = new TreeNode("No results found")
                    {
                        Name = "No results",
                        Text = "No Result"
                    };
                    this.tsResults.Nodes.Add(node2);
                    this.tsResults.CheckBoxes = false;
                }
        }

        /// <summary>
        /// Refactored from a horrible nest of spaghetti code. I don't yet understand this.
        /// </summary>
        /// <param name="searchText"></param>
        /// <param name="usString"></param>
        /// <param name="str2"></param>
        /// <param name="str3"></param>
        /// <param name="moduleName"></param>
        /// <param name="fields"></param>
        /// <returns></returns>
        private static string ConstructQueryTextForModuleName(string searchText, string usString, string str2, string str3, string moduleName, string[] fields)
        {
            string queryText;
            switch (moduleName)
            {
                case "Contacts":
                    queryText = "(contacts.first_name LIKE '%" + clsGlobals.MySqlEscape(usString) + "%' " + str3 + " contacts.last_name LIKE '%" + clsGlobals.MySqlEscape(str2) + "%') OR (contacts.id in (select eabr.bean_id from email_addr_bean_rel eabr INNER JOIN email_addresses ea on eabr.email_address_id = ea.id where eabr.bean_module = 'Contacts' and ea.email_address LIKE '%" + clsGlobals.MySqlEscape(searchText) + "%'))";
                    fields[4] = "account_name";
                    break;
                case "Leads":
                    queryText = "(leads.first_name LIKE '%" + clsGlobals.MySqlEscape(usString) + "%' " + str3 + " leads.last_name LIKE '%" + clsGlobals.MySqlEscape(str2) + "%')  OR (leads.id in (select eabr.bean_id from email_addr_bean_rel eabr INNER JOIN email_addresses ea on eabr.email_address_id = ea.id where eabr.bean_module = 'Leads' and ea.email_address LIKE '%" + clsGlobals.MySqlEscape(searchText) + "%'))";
                    fields[4] = "account_name";
                    break;
                case "Cases":
                    queryText = "(cases.name LIKE '%" + clsGlobals.MySqlEscape(searchText) + "%' OR cases.case_number LIKE '" + clsGlobals.MySqlEscape(searchText) + "')";
                    fields[4] = "case_number";
                    break;
                case "Bugs":
                    queryText = "(bugs.name LIKE '%" + clsGlobals.MySqlEscape(searchText) + "%' " + str3 + " bugs.bug_number LIKE '" + clsGlobals.MySqlEscape(searchText) + "')";
                    fields[4] = "bug_number";
                    break;
                case "Accounts":
                    queryText = "(accounts.name LIKE '%" + clsGlobals.MySqlEscape(usString) + "%') OR (accounts.id in (select eabr.bean_id from email_addr_bean_rel eabr INNER JOIN email_addresses ea on eabr.email_address_id = ea.id where eabr.bean_module = 'Accounts' and ea.email_address LIKE '%" + clsGlobals.MySqlEscape(searchText) + "%'))";
                    fields[4] = "account_name";
                    break;
                default:
                    queryText = moduleName.ToLower() + ".name LIKE '%" + clsGlobals.MySqlEscape(searchText) + "%'";
                    break;
            }

            return queryText;
        }

        private void populateTree(eGetEntryListResult search_result, string module, TreeNode root_node)
        {
            foreach (eEntryValue _value in search_result.entry_list)
            {
                string s = string.Empty;
                string key = string.Empty;
                string valueByKey = string.Empty;
                key = clsSuiteCRMHelper.GetValueByKey(_value, "id");
                s = clsSuiteCRMHelper.GetValueByKey(_value, "first_name") + " " + clsSuiteCRMHelper.GetValueByKey(_value, "last_name");
                if (s == " ")
                {
                    s = clsSuiteCRMHelper.GetValueByKey(_value, "name");
                }
                string str4 = module;
                if (str4 != null)
                {
                    if (!(str4 == "Contacts") && !(str4 == "Leads"))
                    {
                        if (str4 == "Cases")
                        {
                            goto Label_00DC;
                        }
                        if (str4 == "Bugs")
                        {
                            goto Label_00F0;
                        }
                    }
                    else
                    {
                        valueByKey = clsSuiteCRMHelper.GetValueByKey(_value, "account_name");
                    }
                }
                goto Label_0102;
            Label_00DC:
                valueByKey = clsSuiteCRMHelper.GetValueByKey(_value, "case_number");
                goto Label_0102;
            Label_00F0:
                valueByKey = clsSuiteCRMHelper.GetValueByKey(_value, "bug_number");
            Label_0102:
                if (valueByKey != string.Empty)
                {
                    s = s + " (" + valueByKey + ")";
                }
                if (!root_node.Nodes.ContainsKey(key))
                {
                    TreeNode node = new TreeNode(s)
                    {
                        Name = key,
                        Tag = key
                    };
                    root_node.Nodes.Add(node);
                }
            }
            if (search_result.result_count <= 3)
            {
                root_node.Expand();
            }
        }

        private List<CrmEntity> GetSelectedCrmEntities(TreeView tree)
        {
            var result = new List<CrmEntity>();
            foreach (TreeNode node in tree.Nodes)
            {
                this.GetSelectedCrmEntitiesHelper(node, result);
            }
            return result;
        }

        private void GetSelectedCrmEntitiesHelper(TreeNode node, List<CrmEntity> selectedCrmEntities)
        {
            if (((node.Tag != null) && (node.Tag.ToString() != "root_node")) && ((node.Tag.ToString() != "sub_root_node") && node.Checked))
            {
                selectedCrmEntities.Add(new CrmEntity(node.Parent.Text, node.Tag.ToString()));
            }
            foreach (TreeNode node2 in node.Nodes)
            {
                this.GetSelectedCrmEntitiesHelper(node2, selectedCrmEntities);
            }
        }

        private void frmArchive_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                string str = string.Empty;
                for (int i = 0; i < this.lstViewSearchModules.Items.Count; i++)
                {
                    if (this.lstViewSearchModules.Items[i].Checked)
                    {
                        str = str + i + ",";
                    }
                }
                string str2 = str.Remove(str.Length - 1, 1);
                this.settings.SelectedSearchModules = str2;
                this.settings.Save();
                bool flag1 = this.settings.ParticipateInCeip;
            }
            catch (System.Exception)
            {
                // Swallow exception(!)
            }
        }


        private void tsResults_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if ((e.Action != TreeViewAction.Unknown) && (e.Node.Nodes.Count > 0))
            {
                this.CheckAllChildNodes(e.Node, e.Node.Checked);
            }
        }

        private void tsResults_AfterExpand(object sender, TreeViewEventArgs e)
        {
            TreeViewAction action = e.Action;
        }

        private void tsResults_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (((e.Button == MouseButtons.Right) && (e.Node.Tag.ToString() != "root_node")) && (e.Node.Tag.ToString() != "sub_root_node"))
            {
                this.tsResults.SelectedNode = e.Node;
            }
        }

        private void lstViewSearchModules_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            if (e.Item.Text == "All")
            {
                if (e.Item.Checked)
                {
                    for (int i = 1; i < this.lstViewSearchModules.Items.Count; i++)
                    {
                        this.lstViewSearchModules.Items[i].Checked = true;
                    }
                }
                else
                {
                    for (int j = 1; j < this.lstViewSearchModules.Items.Count; j++)
                    {
                        this.lstViewSearchModules.Items[j].Checked = false;
                    }
                }
            }
        }

        private void txtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                this.btnSearch_Click(null, null);
            }
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            if (this.txtSearch.Text == string.Empty)
            {
                this.btnSearch.Enabled = false;
            }
            else
            {
                this.btnSearch.Enabled = true;
            }

        }

        private void btnArchive_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.tsResults.Nodes.Count == 0)
                {
                    MessageBox.Show("There are no search results.", "Error");
                    return;
                }

                var selectedCrmEntities = GetSelectedCrmEntities(this.tsResults);
                if (!selectedCrmEntities.Any())
                {
                    MessageBox.Show("No selected CRM entities", "Error");
                    return;
                }

                var selectedEmailsCount = Globals.ThisAddIn.SelectedEmailCount;
                if (selectedEmailsCount == 0)
                {
                    MessageBox.Show("No emails selected", "Error");
                    return;
                }

                DaemonWorker.Instance.AddTask(new EmailArchiveAction(Globals.ThisAddIn.SelectedEmails, selectedCrmEntities, this.type));

                Close();
            }
            catch (System.Exception exception)
            {
                Log.Error("btnArchive_Click", exception);
                MessageBox.Show("There was an error while archiving", "Error");
            }
        }

        private bool ReportOnEmailArchiveSuccess(List<ArchiveResult> emailArchiveResults)
        {
            var successCount = emailArchiveResults.Count(r => r.IsSuccess);
            var failCount = emailArchiveResults.Count - successCount;
            var fullSuccess = failCount == 0;
            if (fullSuccess)
            {
                if (settings.ShowConfirmationMessageArchive)
                {
                    MessageBox.Show(
                        $"{successCount} email(s) have been successfully archived",
                        "Success");
                }
                return true;
            }
            else
            {
                var message = successCount == 0
                    ? $"Failed to archive {failCount} email(s)"
                    : $"{successCount} emails(s) were successfully archived.";

                var first11Problems = emailArchiveResults.SelectMany(r => r.Problems).Take(11).ToList();
                if (first11Problems.Any())
                {
                    message =
                        message +
                        "\n\nThere were some failures:\n" +
                        string.Join("\n", first11Problems.Take(10)) +
                        (first11Problems.Count > 10 ? "\n[and more]" : "");
                }

                MessageBox.Show(message, "Failure");
                return false;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void txtSearch_Enter(object sender, EventArgs e)
        {
            if (txtSearch.Focused == true)
            {
                this.AcceptButton = btnSearch;
            }
        }

        private void txtSearch_Leave(object sender, EventArgs e)
        {
            this.AcceptButton = btnArchive;
        }
    }
}
