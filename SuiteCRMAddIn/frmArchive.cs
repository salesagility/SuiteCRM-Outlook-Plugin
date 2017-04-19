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
    using Exceptions;
    using Microsoft.Office.Interop.Outlook;
    using SuiteCRMClient;
    using SuiteCRMClient.Email;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Data;
    using System.Linq;
    using System.Text;
    using System.Windows.Forms;
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
               //  StringEnumerator enumerator = .GetEnumerator();
                foreach (string key in this.settings.CustomModules)
                {
                    string[] strArray = key.Split(new char[] { '|' });
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
            using (new WaitCursor(this))
            {
                this.AddActionHandlers();
                this.PopulateUIComponents();

                if (this.settings.AutomaticSearch)
                {
                    this.btnSearch_Click(null, null);
                }
            }
        }

        /// <summary>
        /// Populate my menus and other user interface components.
        /// </summary>
        private void PopulateUIComponents()
        {
            this.txtSearch.Text = ConstructSearchText(Globals.ThisAddIn.SelectedEmails);
            if (this.settings.EmailCategories != null)
            {
                this.categoryInput.DataSource = this.settings.EmailCategories;
            }

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
        }

        /// <summary>
        /// Add handlers for my actions
        /// </summary>
        private void AddActionHandlers()
        {
            this.tsResults.AfterCheck += new TreeViewEventHandler(this.tsResults_AfterCheck);
            this.tsResults.AfterExpand += new TreeViewEventHandler(this.tsResults_AfterExpand);
            this.tsResults.NodeMouseClick += new TreeNodeMouseClickEventHandler(this.tsResults_NodeMouseClick);
            this.txtSearch.KeyDown += new KeyEventHandler(this.txtSearch_KeyDown);
            this.lstViewSearchModules.ItemChecked += new ItemCheckedEventHandler(this.lstViewSearchModules_ItemChecked);
            base.FormClosed += new FormClosedEventHandler(this.frmArchive_FormClosed);
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

            if (this.txtSearch.Text.Replace(';', ',').Contains<char>(','))
            {
                foreach (string str in this.txtSearch.Text.Split(new char[] { ',' }).OrderBy(x => x).GroupBy(x => x).Select(g => g.First()))
                {
                    this.Search(str);
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
            {
                this.txtSearch.Enabled = false;

                try
                {
                    List<string> list = new List<string> { "Accounts", ContactSyncing.CrmModule, "Leads", "Bugs", "Projects", "Cases", "Opportunties" };
                    this.tsResults.CheckBoxes = true;
                    if (searchText == string.Empty)
                    {
                        MessageBox.Show("Please enter some text to search", "Invalid search", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        searchText = searchText.TrimStart(new char[0]);

                        foreach (ListViewItem item in this.lstViewSearchModules.Items)
                        {
                            TreeNode node = null;
                            try
                            {
                                if (item.Checked)
                                {
                                    string moduleName = item.Tag.ToString();

                                    if (moduleName != "All")
                                    {
                                        node = FindOrCreateNodeForModule(moduleName);

                                        List<string> fieldsToSeek = new List<string>();
                                        fieldsToSeek.Add("id");

                                        var queryResult = TryQuery(searchText, moduleName, fieldsToSeek);
                                        if (queryResult != null)
                                        {
                                            if (queryResult.result_count > 0)
                                            {
                                                this.PopulateTree(queryResult, moduleName, node);
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show(
                                                $"An error was encountered while querying module '{moduleName}'. The error has been logged",
                                                $"Query error in module {moduleName}",
                                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        }
                                    }
                                }
                            }
                            catch (System.Exception any)
                            {
                                Globals.ThisAddIn.Log.Error("Failure when custom module included (3)", any);

                                MessageBox.Show(
                                    $"An error was encountered while querying module '{item.Tag.ToString()}'. The error ('{any.Message}') has been logged",
                                    $"Query error in module {item.Tag.ToString()}",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                            finally
                            {
                                if (node != null && node.GetNodeCount(true) <= 0)
                                {
                                    node.Remove();
                                }
                            }
                        }
                    }
                }
                catch (System.Exception any)
                {
#if DEBUG
                    Log.Error("frmArchive.Search: Unexpected error while populating archive tree", any);
#endif
                    this.tsResults.Nodes.Clear();
                }
                finally
                {
                    if (this.tsResults.Nodes.Count <= 0)
                    {
                        ShowNoResults(this.tsResults);
                    }
                    else
                    {
                        this.btnArchive.Enabled = true;
                    }
                    this.txtSearch.Enabled = true;
                }
            }
        }

        /// <summary>
        /// Clear out this tree and show a single unselectable node labelled 'No results found'
        /// </summary>
        private void ShowNoResults(TreeView tree)
        {
            tree.Nodes.Clear();
            TreeNode node2 = new TreeNode("No results found")
            {
                Name = "No results",
                Text = "No Result"
            };
            tree.Nodes.Add(node2);
            tree.CheckBoxes = false;
        }

        private eGetEntryListResult TryQuery(string searchText, string moduleName, List<string> fieldsToSeek)
        {
            eGetEntryListResult queryResult = null;

            string queryText = ConstructQueryTextForModuleName(searchText, moduleName, fieldsToSeek);
            try
            {
                queryResult = clsSuiteCRMHelper.GetEntryList(moduleName, queryText, settings.SyncMaxRecords, "date_entered DESC", 0, false, fieldsToSeek.ToArray());
            }
            catch (System.Exception any)
            {
                Globals.ThisAddIn.Log.Error($"Failure when custom module included (1)\n\tQuery was '{queryText}'", any);

                queryText = queryText.Replace("%", string.Empty);
                try
                {
                    queryResult = clsSuiteCRMHelper.GetEntryList(moduleName, queryText, settings.SyncMaxRecords, "date_entered DESC", 0, false, fieldsToSeek.ToArray());
                }
                catch (Exception secondFail)
                {
                    Globals.ThisAddIn.Log.Error($"Failure when custom module included (2)\n\tQuery was '{queryText}'", secondFail);
                    queryResult = null;
                    throw;
                }
                if (queryResult == null)
                {
                    throw;
                }
            }

            return queryResult;
        }

        /// <summary>
        /// Find the existing node in my tsResults tree view which represents the module with
        /// this name; if none is found, create one and add it to the tree.
        /// </summary>
        /// <param name="moduleName">The name of the module to be found.</param>
        /// <returns>A suitable tree node.</returns>
        private TreeNode FindOrCreateNodeForModule(string moduleName)
        {
            TreeNode node;
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

            return node;
        }

        /// <summary>
        /// Construct suitable query text to query the module with this name for potential connection with this search text.
        /// </summary>
        /// <remarks>
        /// Refactored from a horrible nest of spaghetti code. I don't yet fully understand this.
        /// </remarks>
        /// <param name="searchText">The text entered to search for.</param>
        /// <param name="moduleName">The name of the module to search.</param>
        /// <param name="fields">The list of fields to pull back, which may be modified by this method.</param>
        /// <returns>A suitable search query</returns>
        /// <exception cref="CouldNotConstructQueryException">if no search string could be constructed.</exception>
        private static string ConstructQueryTextForModuleName(string searchText, string moduleName, List<string> fields)
        {
            string queryText = string.Empty;
            List<string> searchTokens = searchText.Split(new char[] { ' ' }).ToList();
            var escapedSearchText = clsGlobals.MySqlEscape(searchText);
            var firstTerm = clsGlobals.MySqlEscape(searchTokens.First());
            var lastTerm = clsGlobals.MySqlEscape(searchTokens.Last());
            string logicalOperator = firstTerm == lastTerm ? "OR" : "AND";

            switch (moduleName)
            {
                case ContactSyncing.CrmModule:
                    queryText = ConstructQueryTextWithFirstAndLastNames(moduleName, escapedSearchText, firstTerm, lastTerm);
                    AddFirstLastAndAccountNames(fields);
                    break;
                case "Leads":
                    queryText = ConstructQueryTextWithFirstAndLastNames(moduleName, escapedSearchText, firstTerm, lastTerm);
                    AddFirstLastAndAccountNames(fields);
                    break;
                case "Cases":
                    queryText = $"(cases.name LIKE '%{escapedSearchText}%' OR cases.case_number LIKE '{escapedSearchText}')";
                    foreach (string fieldName in new string[] { "name", "case_number" })
                    {
                        fields.Add(fieldName);
                    }
                    break;
                case "Bugs":
                    queryText = $"(bugs.name LIKE '%{escapedSearchText}%' {logicalOperator} bugs.bug_number LIKE '{escapedSearchText}')";
                    foreach (string fieldName in new string[] { "name", "bug_number" })
                    {
                        fields.Add(fieldName);
                    }
                    break;
                case "Accounts":
                    queryText = "(accounts.name LIKE '%" + firstTerm + "%') OR (accounts.id in (select eabr.bean_id from email_addr_bean_rel eabr INNER JOIN email_addresses ea on eabr.email_address_id = ea.id where eabr.bean_module = 'Accounts' and ea.email_address LIKE '%" + escapedSearchText + "%'))";
                    AddFirstLastAndAccountNames(fields);
                    break;
                default:
                    List<string> fieldNames = clsSuiteCRMHelper.GetCharacterFields(moduleName);

                    if (fieldNames.Contains("first_name") && fieldNames.Contains("last_name"))
                    {
                        /* This is Ian's suggestion */
                        queryText = ConstructQueryTextWithFirstAndLastNames(moduleName, escapedSearchText, firstTerm, lastTerm);
                        foreach (string fieldName in new string[] { "first_name", "last_name" })
                        {
                            fields.Add(fieldName);
                        }
                    }
                    else
                    {
                        queryText = ConstructQueryTextForUnknownModule(moduleName, escapedSearchText);
                    }
                    break;
            }

            if (string.IsNullOrEmpty(queryText))
            {
                throw new CouldNotConstructQueryException(moduleName, searchText);
            }

            return queryText;
        }

        /// <summary>
        /// Add 'first_name', 'last_name', 'name' and 'account_name' to this list of field names.
        /// </summary>
        /// <remarks>
        /// It feels completely wrong to do this in code. There should be a configuration file of
        /// appropriate fieldnames for module names somewhere, but there isn't. TODO: probably fix.
        /// </remarks>
        /// <param name="fields">The list of field names.</param>
        private static void AddFirstLastAndAccountNames(List<string> fields)
        {
            foreach (string fieldName in new string[] { "first_name", "last_name", "name", "account_name" })
            {
                fields.Add(fieldName);
            }
        }

        /// <summary>
        /// If we really don't know anything about the module (which is likely with custom modules)
        /// the best we can do is see whether any of its text fields matches the search text.
        /// </summary>
        /// <param name="moduleName">The name of the module to search.</param>
        /// <param name="escapedSearchText">The text to search for, escaped for MySQL.</param>
        /// <returns>A query string.</returns>
        private static string ConstructQueryTextForUnknownModule(string moduleName, string escapedSearchText)
        {
            StringBuilder queryBuilder = new StringBuilder();

            if (clsSuiteCRMHelper.GetActivitiesLinks(moduleName, Objective.Email).Count() > 0)
            {
                string tableName = moduleName.ToLower();

                foreach (string fieldName in clsSuiteCRMHelper.GetCharacterFields(moduleName))
                {
                    if (!String.IsNullOrWhiteSpace(queryBuilder.ToString()))
                    {
                        queryBuilder.Append("OR ");
                    }
                    queryBuilder.Append($"{tableName}.{fieldName} LIKE '%{escapedSearchText}%' ");
                }

                queryBuilder.Append($"OR {tableName} in (select eabr.bean_id from email_addr_bean_rel eabr ")
                    .Append("INNER JOIN email_addresses ea on eabr.email_address_id = ea.id ")
                    .Append($"where eabr.bean_module = '{moduleName}' ")
                    .Append($" and ea.email_address LIKE '{escapedSearchText}'");
            }

            return queryBuilder.ToString();
        }

        /// <summary>
        /// If the search text supplied comprises two space-separated tokens, these are possibly a first and last name;
        /// if only one token, it's likely an email address, but might be a first or last name. 
        /// This query explores those possibilities.
        /// </summary>
        /// <param name="moduleName">The name of the module to search.</param>
        /// <param name="emailAddress">The portion of the search text which may be an email address.</param>
        /// <param name="firstName">The portion of the search text which may be a first name</param>
        /// <param name="lastName">The portion of the search text which may be a last name</param>
        /// <returns>If the module has fields 'first_name' and 'last_name', then a potential query string;
        /// else an empty string.</returns>
        private static string ConstructQueryTextWithFirstAndLastNames(
            string moduleName, 
            string emailAddress, 
            string firstName, 
            string lastName)
        {
            List<string> fieldNames = clsSuiteCRMHelper.GetCharacterFields(moduleName);
            string result = string.Empty;
            string logicalOperator = firstName == lastName ? "OR" : "AND";

            if (fieldNames.Contains("first_name") && fieldNames.Contains("last_name"))
            {
                string lowerName = moduleName.ToLower();
                result = $"({lowerName}.first_name LIKE '%{firstName}%' {logicalOperator} {lowerName}.last_name LIKE '%{lastName}%') OR ({lowerName}.id in (select eabr.bean_id from email_addr_bean_rel eabr INNER JOIN email_addresses ea on eabr.email_address_id = ea.id where eabr.bean_module = '{moduleName}' and ea.email_address LIKE '%{emailAddress}%'))"; ;
            }

            return result;
        }

        /// <summary>
        /// Add a node beneath this parent representing this search result in this module.
        /// </summary>
        /// <param name="searchResult">A search result</param>
        /// <param name="module">The module in which the search was performed</param>
        /// <param name="parent">The parent node beneath which the new node should be added.</param>
        private void PopulateTree(eGetEntryListResult searchResult, string module, TreeNode parent)
        {
            foreach (eEntryValue entry in searchResult.entry_list)
            {
                string key = clsSuiteCRMHelper.GetValueByKey(entry, "id");
                if (!parent.Nodes.ContainsKey(key))
                {
                    TreeNode node = new TreeNode(ConstructNodeName(module, entry))
                    {
                        Name = key,
                        Tag = key
                    };
                    parent.Nodes.Add(node);
                }
            }
            if (searchResult.result_count <= 3)
            {
                parent.Expand();
            }
        }

        /// <summary>
        /// Construct suitable label text for a tree node representing this value in this module.
        /// </summary>
        /// <param name="module">The name of the module.</param>
        /// <param name="entry">The value in the module.</param>
        /// <returns>A canonical tree node label.</returns>
        private static string ConstructNodeName(string module, eEntryValue entry)
        {
            StringBuilder nodeNameBuilder = new StringBuilder();
            string keyValue = string.Empty;
            nodeNameBuilder.Append(clsSuiteCRMHelper.GetValueByKey(entry, "first_name"))
                .Append(" ")
                .Append(clsSuiteCRMHelper.GetValueByKey(entry, "last_name"));

            if (String.IsNullOrWhiteSpace(nodeNameBuilder.ToString()))
            {
                nodeNameBuilder.Append(clsSuiteCRMHelper.GetValueByKey(entry, "name"));
            }

            switch (module)
            {
                case "Cases":
                    keyValue = clsSuiteCRMHelper.GetValueByKey(entry, "case_number");
                    break;
                case "Bugs":
                    keyValue = clsSuiteCRMHelper.GetValueByKey(entry, "bug_number");
                    break;
                default:
                    keyValue = clsSuiteCRMHelper.GetValueByKey(entry, "account_name");
                    break;
            }

            if (keyValue != string.Empty)
            {
                nodeNameBuilder.Append($" ({keyValue})");
            }

            return nodeNameBuilder.ToString();
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

                var archiver = new EmailArchiving($"EB-{Globals.ThisAddIn.SelectedEmailCount}", Globals.ThisAddIn.Log);
                this.ReportOnEmailArchiveSuccess(
                    Globals.ThisAddIn.SelectedEmails.Select(mailItem =>
                            archiver.ArchiveEmailWithEntityRelationships(mailItem, selectedCrmEntities, this.type))
                        .ToList());

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
                        (first11Problems.Count > 10 ? "\n[and more]" : string.Empty);
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
