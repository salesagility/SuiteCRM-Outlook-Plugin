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
    using Exceptions;
    using Extensions;
    using Microsoft.Office.Interop.Outlook;
    using Newtonsoft.Json.Linq;
    using SuiteCRMClient;
    using SuiteCRMClient.Email;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Windows.Forms;
    using Exception = System.Exception;

    public partial class ArchiveDialog : Form
    {
        public readonly List<string> standardModules = new List<string> { "Accounts", "Bugs", "Cases", ContactSynchroniser.CrmModule, "Leads", "Opportunities", "Project", "Users" };

        /// <summary>
        /// The emails to be archived.
        /// </summary>
        private readonly IEnumerable<MailItem> archivableEmails;

        private ILogger Log => Globals.ThisAddIn.Log;

        /// <summary>
        /// Chains of modules to expand in the search tree.
        /// </summary>
        private IDictionary<string, ICollection<LinkSpec>> moduleChains = new Dictionary<string, ICollection<LinkSpec>>();

        /// <summary>
        /// The reason for archiving.
        /// </summary>
        private readonly EmailArchiveReason reason;

        public ArchiveDialog(IEnumerable<MailItem> selectedEmails, EmailArchiveReason reason)
        {
            InitializeComponent();

            try
            {
                using (WaitCursor.For(this))
                {
                    this.moduleChains = AdvancedArchiveSettingsDialog.SetupSearchChains();
                }
            }
            catch (Exception any)
            {
                ErrorHandler.Handle($"Failed to parse Archiving Search Chains value '{Properties.Settings.Default.ArchivingSearchChains}':", any);
            }

            this.archivableEmails = selectedEmails;
            this.reason = reason;

            var alreadyArchived = selectedEmails.Where(x => !string.IsNullOrEmpty(x.GetCRMEntryId()));
            var anyArchived = alreadyArchived.Any();
            this.legend.Text = anyArchived ?
                $"{alreadyArchived.Count()} message(s) have already been archived; rearchiving will remove any existing relationships. You must select all contacts, accounts, leads, etc that you wish to archive the message(s) to." :
                "";
            this.legend.Visible = anyArchived;

            if (!anyArchived)
            {
                this.tsResults.Height += this.legend.Height;
                this.lstViewSearchModules.Height += this.legend.Height;
            }
        }


        /// <summary>
        /// Add any selected custom modules to the list view
        /// </summary>
        private void AddCustomModules()
        {
            if (Properties.Settings.Default.CustomModules != null)
            {
               //  StringEnumerator enumerator = .GetEnumerator();
                foreach (string key in Properties.Settings.Default.CustomModules)
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

        /// <summary>
        /// Add the standard modules to the list view.
        /// </summary>
        private void AddStandardModules()
        {
            AvailableModules allModules = RestAPIWrapper.GetModules();
            foreach (string moduleKey in this.standardModules.OrderBy(x => x))
            {
                var module = allModules.items.FirstOrDefault(x => x.module_key == moduleKey);
                if (module != null)
                {
                    this.lstViewSearchModules.Items.Add(new ListViewItem
                    {
                        Tag = module.module_key,
                        Text = module.module_label
                    });
                }
                else
                {
                    Log.Warn($"Standard modules '{moduleKey}' was not found on the CRM system");
                }
            }
        }

        private void frmArchive_Load(object sender, EventArgs e)
        {
            this.AddActionHandlers();
            this.PopulateUIComponents();

            if (Properties.Settings.Default.AutomaticSearch)
            {
               this.BeginInvoke((MethodInvoker)delegate
               {
                   this.Search(this.txtSearch.Text);
               });
            }
        }

        /// <summary>
        /// Populate my menus and other user interface components.
        /// </summary>
        private void PopulateUIComponents()
        {
            this.txtSearch.Text = ConstructSearchText();
            if (Properties.Settings.Default.EmailCategories != null && Properties.Settings.Default.EmailCategories.IsImplemented)
            {
                this.categoryInput.DataSource = Properties.Settings.Default.EmailCategories;
            }
            else
            {
                this.categoryInput.Enabled = false;
                this.categoryInput.Visible = false;
                this.categoryLabel.Visible = false;
            }

            this.AddStandardModules();

            if (Properties.Settings.Default.ShowCustomModules)
            {
                this.AddCustomModules();
            }
            try
            {
                foreach (string str in Properties.Settings.Default.SelectedSearchModules.Split(new char[] { ',' }))
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
        private string ConstructSearchText()
        {
            string result;
            var currentUserSMTPAddress = Globals.ThisAddIn.Application.GetCurrentUserSMTPAddress();
            switch (this.reason)
            {
                case EmailArchiveReason.Inbound:
                    result = ConstructInboundSearchTest();
                    break;
                case EmailArchiveReason.Manual:
                    // is the sender the current user?
                    if (this.archivableEmails
                        .Select(email => email.GetSenderSMTPAddress())
                        .Any(x => x.Equals(currentUserSMTPAddress)))
                    {
                        if (this.archivableEmails
                            .Select(email => email.GetSenderSMTPAddress())
                            .All(x => x.Equals(currentUserSMTPAddress)))
                        {
                            result = ConstructOutboundSearchText();
                        }
                        else
                        {
                            ICollection<string> addresses = 
                                new HashSet<string>($"{ConstructInboundSearchTest()},{ConstructOutboundSearchText()}"
                                .Split(','));
                            addresses.Remove(currentUserSMTPAddress);
                            result = string.Join(",", addresses.OrderBy(address => address)
                                        .GroupBy(address => address)
                                        .Select(g => g.First()));
                        }
                        }
                    else
                    {
                        result = ConstructInboundSearchTest();
                    }
                    break;
                case EmailArchiveReason.Outbound:
                case EmailArchiveReason.SendAndArchive:
                    result = ConstructOutboundSearchText();
                    break;
                default:
                    result = string.Empty;
                    break;
            }

            return result;
        }

        private string ConstructOutboundSearchText()
        {
            string result;
            ICollection<string> addresses = new HashSet<string>();
            foreach (var email in this.archivableEmails)
            {
                foreach (Recipient recipient in email.Recipients)
                    addresses.Add(recipient.GetSmtpAddress());
            }
            result = string.Join(",", addresses.OrderBy(address => address)
                        .GroupBy(address => address)
                        .Select(g => g.First()));
            return result;
        }

        private string ConstructInboundSearchTest()
        {
            return string.Join(",", this.archivableEmails.Select(email => email.GetSenderSMTPAddress())
                        .OrderBy(address => address)
                        .GroupBy(address => address)
                        .Select(g => g.First()));
        }


        /// <summary>
        /// Set the checkboxes of all children of this node to this value.
        /// </summary>
        /// <param name="node">The parent of the nodes to change.</param>
        /// <param name="value">The value to set.</param>
        private void CheckAllChildNodes(TreeNode node, bool value)
        {
            foreach (TreeNode child in node.Nodes)
            {
                child.Checked = value;
                this.CheckAllChildNodes(child, value);
            }
        }

        public void btnSearch_Click(object sender, EventArgs e)
        {
            this.tsResults.Nodes.Clear();

            this.Search(this.txtSearch.Text);

            this.AcceptButton = btnArchive;
        }

        private bool UnallowedNumber(string strText)
        {
            return Char.IsDigit(strText.First());
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            base.Close();
        }


        /// <summary>
        /// Search CRM for records matching this search text, and populate the tree view
        /// with a tree of nodes representing the records found.
        /// </summary>
        /// <remarks>
        /// TODO: Candidate for refactoring.
        /// </remarks>
        /// <param name="allSearchText">The text to search for.</param>
        public void Search(string allSearchText)
        {
            this.txtSearch.Enabled = false;

            foreach (string searchText in
                allSearchText.Split(new char[] { ',', ';' })
                .OrderBy(x => x)
                .GroupBy(x => x).Select(g => g.First().Trim()))
            {

                try
                {
                    this.tsResults.CheckBoxes = true;
                    if (searchText == string.Empty)
                    {
                        MessageBox.Show("Please enter some text to search", "Invalid search", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
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
                                        EntryList queryResult;
                                        using (WaitCursor.For(this))
                                        {
                                            queryResult = TryQuery(searchText, moduleName, fieldsToSeek);
                                        }
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
                                ErrorHandler.Handle("Failure when custom module included (3)", any);

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
                catch (Exception any)
                {
#if DEBUG
                    ErrorHandler.Handle("Unexpected error while populating archive tree", any);
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
                        this.tsResults.Sort();
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

        private EntryList TryQuery(string searchText, string moduleName, List<string> fieldsToSeek)
        {
            EntryList queryResult = null;

            var searchQuery = new SearchQuery(searchText, moduleName, fieldsToSeek, moduleChains);
            try
            {
                queryResult = searchQuery.Execute();
            }
            catch (System.Exception any)
            {
                ErrorHandler.Handle($"Failure when custom module included (1); Query was '{searchQuery}'", any);

                searchQuery = searchQuery.Replace("%", string.Empty);
                try
                {
                    queryResult = searchQuery.Execute();
                }
                catch (Exception secondFail)
                {
                    ErrorHandler.Handle($"Failure when custom module included (2); Query was '{searchQuery}'", secondFail);
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
        /// If we really don't know anything about the module (which is likely with custom modules)
        /// the best we can do is see whether any of its text fields matches the search text.
        /// </summary>
        /// <param name="moduleName">The name of the module to search.</param>
        /// <param name="escapedSearchText">The text to search for, escaped for MySQL.</param>
        /// <returns>A query string.</returns>
        private static string ConstructQueryTextForUnknownModule(string moduleName, string escapedSearchText)
        {
            StringBuilder queryBuilder = new StringBuilder();

            if (RestAPIWrapper.GetActivitiesLinks(moduleName, Objective.Email).Count() > 0)
            {
                string tableName = moduleName.ToLower();

                foreach (string fieldName in RestAPIWrapper.GetCharacterFields(moduleName))
                {
                    switch (fieldName)
                    {
                        case "name":
                        case "first_name":
                        case "last_name":
                            queryBuilder.Append($"{tableName}.{fieldName} LIKE '%{escapedSearchText}%' OR ");
                            break;
                    }
                }

                queryBuilder.Append($"{tableName}.id in (select eabr.bean_id from email_addr_bean_rel eabr ")
                    .Append("INNER JOIN email_addresses ea on eabr.email_address_id = ea.id ")
                    .Append($"where eabr.bean_module = '{moduleName}' ")
                    .Append($" and ea.email_address LIKE '{escapedSearchText}')");
            }

            return queryBuilder.ToString();
        }


        /// <summary>
        /// Add a node beneath this parent representing this search result in this module.
        /// </summary>
        /// <param name="searchResult">A search result</param>
        /// <param name="module">The module in which the search was performed</param>
        /// <param name="parent">The parent node beneath which the new node should be added.</param>
        private void PopulateTree(EntryList searchResult, string module, TreeNode parent)
        {
            for (int i = 0; i < searchResult.entry_list.Count(); i++)
            {
                EntryValue entry = searchResult.entry_list[i];
                string key = RestAPIWrapper.GetValueByKey(entry, "id");

                ICollection<LinkSpec> links = moduleChains.ContainsKey(module) ? moduleChains[module] : new List<LinkSpec>();

                if (!parent.Nodes.ContainsKey(key))
                {
                    TreeNode node = new TreeNode(ConstructNodeName(module, entry))
                    {
                        Name = key,
                        Tag = key
                    };
                    parent.Nodes.Add(node);

                    foreach (var relationship in entry.relationships.link_list)
                    {
                        LinkSpec link = links.Where(x => x.LinkName == relationship.name).FirstOrDefault();

                        TreeNode chainNode = new TreeNode(link != null ? link.TargetName : relationship.name);
                        node.Nodes.Add(chainNode);

                        foreach (LinkRecord member in relationship.records)
                        {
                            var targetId = link != null ? link.TargetId : "id";

                            TreeNode memberNode = new TreeNode(member.data.GetValueAsString(link != null ? link.Label : "name"))
                            {
                                Name = member.data.GetValueAsString(targetId),
                                Tag = member.data.GetValueAsString(targetId)
                            };
                            chainNode.Nodes.Add(memberNode);
                        }
                    }
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
        private static string ConstructNodeName(string module, EntryValue entry)
        {
            StringBuilder nodeNameBuilder = new StringBuilder();
            string keyValue = string.Empty;
            nodeNameBuilder.Append(RestAPIWrapper.GetValueByKey(entry, "first_name"))
                .Append(" ")
                .Append(RestAPIWrapper.GetValueByKey(entry, "last_name"));

            if (String.IsNullOrWhiteSpace(nodeNameBuilder.ToString()))
            {
                nodeNameBuilder.Append(RestAPIWrapper.GetValueByKey(entry, "name"));
            }

            switch (module)
            {
                case "Bugs":
                    keyValue = RestAPIWrapper.GetValueByKey(entry, "bug_number");
                    break;
                case "Cases":
                    keyValue = RestAPIWrapper.GetValueByKey(entry, "case_number");
                    break;
                case "Contacts":
                    keyValue = string.Empty;
                    break;
                default:
                    keyValue = RestAPIWrapper.GetValueByKey(entry, "account_name");
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
                Properties.Settings.Default.SelectedSearchModules = str2;
                Properties.Settings.Default.Save();
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
            this.AcceptButton = btnSearch;

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
                IEnumerable<MailItem> itemsToArchive = ConfirmRearchiveAlreadyArchivedEmails.ConfirmAlreadyArchivedEmails(archivableEmails);

                using (WaitCursor.For(this))
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

                        var selectedEmailsCount = itemsToArchive.Count();
                        if (selectedEmailsCount > 0)
                        {
                            Log.Debug($"ArchiveDialog: About to manually archive {selectedEmailsCount} emails");
                            var archiver = Globals.ThisAddIn.EmailArchiver;
                            bool success = this.ReportOnEmailArchiveSuccess(
                                itemsToArchive.Select(mailItem =>
                                        archiver.ArchiveEmailWithEntityRelationships(mailItem, selectedCrmEntities, this.reason))
                                    .ToList());
                            string prefix = success ? "S" : "Uns";
                            Log.Debug($"ArchiveDialog: {prefix}uccessfully archived {selectedEmailsCount} emails");

                            this.DialogResult = DialogResult.OK;
                            Close();
                        }
                        else
                        {
                            this.DialogResult = DialogResult.Cancel;
                        }
                    }
                    catch (System.Exception exception)
                    {
                        this.DialogResult = DialogResult.Abort;
                        Log.Error("btnArchive_Click", exception);
                        MessageBox.Show("There was an error while archiving", "Error");
                    }
                    finally
                    {
                        foreach (MailItem i in this.archivableEmails)
                        {
                            i.Save();
                        }
                    }
                }
            }
            catch (DialogCancelledException)
            {

            }
        }


        /// <summary>
        /// Construct a message reporting on the success or failure of archiving represented
        /// by these results, and show it in a message box.
        /// </summary>
        /// <remarks>
        /// TODO: Candidate for significant refactoring - this is ugly.
        /// </remarks>
        /// <param name="results">The email archiving results to report on.</param>
        /// <returns>true if the report indicated success, else false.</returns>
        private bool ReportOnEmailArchiveSuccess(List<ArchiveResult> results)
        {
            var successCount = results.Count(r => r.IsSuccess);
            var failCount = results.Count - successCount;
            var fullSuccess = failCount == 0;
            if (fullSuccess)
            {
                if (Properties.Settings.Default.ShowConfirmationMessageArchive)
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

                var allProblems = results
                    .Where(r => r.Problems != null)
                    .SelectMany(r => r.Problems);

                if (allProblems.Any())
                {
                    message =
                        message +
                        "\n\nThere were some failures:\n" +
                        string.Join("\n", allProblems.Take(10)) +
                        (allProblems.Count() > 10 ? "\n[and more]" : string.Empty);
                }

                MessageBox.Show(message, "Failure");
                return false;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            base.Close();
        }

        private void txtSearch_Enter(object sender, EventArgs e)
        {
            if (txtSearch.Focused == true)
            {
                this.AcceptButton = btnSearch;
            }
        }

        /// <summary>
        /// If the Category input changes, set its current value as the Category property on 
        /// each of the currently selected emails.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void categoryInput_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (MailItem mail in this.archivableEmails)
            {
                try
                {
                    if (mail.UserProperties[MailItemExtensions.CRMCategoryPropertyName] == null)
                    {
                        mail.UserProperties.Add(MailItemExtensions.CRMCategoryPropertyName, OlUserPropertyType.olText);
                    }
                    mail.UserProperties[MailItemExtensions.CRMCategoryPropertyName].Value = categoryInput.SelectedItem.ToString();
                }
                catch (COMException)
                {
                    /* this happens with SendAndArchive. Don't know why, but you can neither access the 
                     * UserProperties nor add to them */
                }
            }
        }
        private void RemoveSelection(Object obj)
        {
            TextBox textbox = obj as TextBox;
            if (textbox != null)
            {
                textbox.SelectionLength = 0;
            }
        }

        private void legend_MouseUp(object sender, MouseEventArgs e)
        {
            RemoveSelection(sender);
        }

        private void legend_KeyUp(object sender, KeyEventArgs e)
        {
            RemoveSelection(sender);
        }

        internal class LinkSpec
        {
            /// <summary>
            /// The name of the linking module
            /// </summary>
            public string LinkName { get; set; }

            /// <summary>
            /// The names of the fields to request
            /// </summary>
            public ICollection<string> FieldNames { get; set; }

            /// <summary>
            /// The name of the target module
            /// </summary>
            public string TargetName { get; set; }

            /// <summary>
            /// The name of the id field in the target module (expected to be one of the fields in `FieldNames`).
            /// </summary>
            public string TargetId { get; set; }

            /// <summary>
            /// The label to put on the node.
            /// </summary>
            public string Label { get; set; }

            public LinkSpec(string linkName, string targetName, ICollection<string> fieldNames)
            {
                this.LinkName = linkName;
                this.TargetName = targetName;
                this.FieldNames = fieldNames;
                this.TargetId = fieldNames.ElementAt(0);
                this.Label = fieldNames.ElementAtOrDefault(1);
            }

        }

        /// <summary>
        /// A query to search the server side database in response to a client 
        /// side (at this state, ArchiveDialog) query; a testable bean.
        /// </summary>
        internal class SearchQuery
        {
            /// <summary>
            /// The text we're seeking.
            /// </summary>
            private string searchText;
            /// <summary>
            /// The name of the module within which we're seeking it.
            /// </summary>
            private string moduleName;
            /// <summary>
            /// The fields within which we're seeking.
            /// </summary>
            private List<string> fieldsToSeek;
            /// <summary>
            /// The chains of modules we should traverse.
            /// </summary>
            private IDictionary<string, ICollection<LinkSpec>> moduleChains;
            /// <summary>
            /// A collection of substitutions to make in the constructed search text.
            /// </summary>
            private IList<Tuple<Regex, string>> replacements = new List<Tuple<Regex, string>>();

            /// <summary>
            /// A collection of modulename, fieldname pairs to order by.
            /// </summary>
            private ISet<Tuple<string,string>> fieldsToOrderBy = new HashSet<Tuple<string,string>>();

 
            /// <summary>
            /// An arguably redundant property wrapped around 
            /// ConstructSearchText, mainly for tesstability.
            /// </summary>
            public string QueryText { get
                {
                    return this.ConstructQueryText();
                }
            }

            /// <summary>
            /// Get the text of a clause to order the quwery results by.
            /// </summary>
            /// <remarks>Note that the `order_by` clause of `get_entry_list` 
            /// doesn't work reliably, because of deduping.</remarks>
            public string OrderClause { get
                {
                    return this.fieldsToOrderBy.Any() ? 
                        String.Join(",", this.fieldsToOrderBy.Select(x=> x.Item2)) : 
                        "date_entered DESC";
                }
            }

            /// <summary>
            /// Construct a new instance of a search query.
            /// </summary>
            /// <param name="searchText">The text entered to search for.</param>
            /// <param name="moduleName">The name of the module to search.</param>
            /// <param name="fieldsToSeek">The list of fields to pull back, which may be modified by this method.</param>
            public SearchQuery(string searchText, string moduleName, List<string> fieldsToSeek, IDictionary<string, ICollection<LinkSpec>> moduleChains)
            {
                this.searchText = searchText;
                this.moduleName = moduleName;
                this.fieldsToSeek = fieldsToSeek;
                this.moduleChains = moduleChains;
            }

            /// <summary>
            /// If the search text supplied comprises two space-separated tokens, these are possibly a first and last name;
            /// if only one token, it's likely an email address, but might be a first or last name. 
            /// This query explores those possibilities.
            /// </summary>
            /// <param name="emailAddress">The portion of the search text which may be an email address.</param>
            /// <param name="firstName">The portion of the search text which may be a first name</param>
            /// <param name="lastName">The portion of the search text which may be a last name</param>
            /// <returns>If the module has fields 'first_name' and 'last_name', then a potential query string;
            /// else an empty string.</returns>
            private string ConstructQueryTextWithFirstAndLastNames(
                string emailAddress,
                string firstName,
                string lastName)
            {
                List<string> fieldNames = RestAPIWrapper.GetCharacterFields(this.moduleName);
                string result = string.Empty;
                string logicalOperator = firstName == lastName ? "OR" : "AND";

                if (fieldNames.Contains("first_name") && fieldNames.Contains("last_name"))
                {
                    string moduleLower = this.moduleName.ToLower();
                    fieldsToOrderBy.Add(new Tuple<string, string>(moduleLower, "last_name"));
                    fieldsToOrderBy.Add(new Tuple<string, string>(moduleLower, "first_name"));
                    fieldsToOrderBy.Add(new Tuple<string, string>("ea", "email_address"));

                    result = $"({moduleLower}.first_name LIKE '%{firstName}%' {logicalOperator} {moduleLower}.last_name LIKE '%{lastName}%') OR ({moduleLower}.id in (select eabr.bean_id from email_addr_bean_rel eabr INNER JOIN email_addresses ea on eabr.email_address_id = ea.id where eabr.bean_module = '{moduleName}' and ea.email_address LIKE '%{emailAddress}%'))"; ;
                }

                return result;
            }


            /// <summary>
            /// Add 'first_name', 'last_name', 'name' and 'account_name' to this list of field names.
            /// </summary>
            /// <remarks>
            /// It feels completely wrong to do this in code. There should be a configuration file of
            /// appropriate fieldnames for module names somewhere, but there isn't. TODO: probably fix.
            /// </remarks>
            /// <param name="fields">The list of field names.</param>
            private void AddFirstLastAndAccountNames(List<string> fields)
            {
                foreach (string fieldName in new string[] { "first_name", "last_name", "name", "account_name" })
                {
                    fields.Add(fieldName);
                    this.fieldsToOrderBy.Add(new Tuple<string, string>(this.moduleName, fieldName));
                }
            }


            /// <summary>
            /// Construct suitable query text to query the module with this name for potential connection with this search text.
            /// </summary>
            /// <remarks>
            /// Refactored from a horrible nest of spaghetti code. I don't yet fully understand this.
            /// TODO: Candidate for further refactoring to reduce complexity.
            /// </remarks>
            /// <returns>A suitable search query</returns>
            /// <exception cref="CouldNotConstructQueryException">if no search string could be constructed.</exception>
            private string ConstructQueryText()
            {
                string queryText = string.Empty;
                List<string> searchTokens = searchText.Split(new char[] { ',', ';' }).ToList();
                var escapedSearchText = RestAPIWrapper.MySqlEscape(searchText);
                var firstTerm = RestAPIWrapper.MySqlEscape(searchTokens.First());
                var lastTerm = RestAPIWrapper.MySqlEscape(searchTokens.Last());
                string logicalOperator = firstTerm == lastTerm ? "OR" : "AND";

                switch (moduleName)
                {
                    case ContactSynchroniser.CrmModule:
                        queryText = ConstructQueryTextWithFirstAndLastNames(escapedSearchText, firstTerm, lastTerm);
                        AddFirstLastAndAccountNames(fieldsToSeek);
                        break;
                    case "Leads":
                        queryText = ConstructQueryTextWithFirstAndLastNames(escapedSearchText, firstTerm, lastTerm);
                        AddFirstLastAndAccountNames(fieldsToSeek);
                        break;
                    case "Cases":
                        queryText = $"(cases.name LIKE '%{escapedSearchText}%' OR cases.case_number LIKE '{escapedSearchText}')";
                        foreach (string fieldName in new string[] { "name", "case_number" })
                        {
                            fieldsToSeek.Add(fieldName);
                            fieldsToOrderBy.Add(new Tuple<string, string>(this.moduleName, fieldName));
                        }
                        break;
                    case "Bugs":
                        queryText = $"(bugs.name LIKE '%{escapedSearchText}%' {logicalOperator} bugs.bug_number LIKE '{escapedSearchText}')";
                        foreach (string fieldName in new string[] { "name", "bug_number" })
                        {
                            fieldsToSeek.Add(fieldName);
                            fieldsToOrderBy.Add(new Tuple<string, string>(this.moduleName, fieldName));
                        }
                        break;
                    case "Accounts":
                        queryText = "(accounts.name LIKE '%" + firstTerm + "%') OR (accounts.id in (select eabr.bean_id from email_addr_bean_rel eabr INNER JOIN email_addresses ea on eabr.email_address_id = ea.id where eabr.bean_module = 'Accounts' and ea.email_address LIKE '%" + escapedSearchText + "%'))";
                        foreach (string fieldName in new string[] { "name", "account_name" })
                        {
                            this.fieldsToSeek.Add(fieldName);
                            this.fieldsToOrderBy.Add(new Tuple<string, string>(this.moduleName, fieldName));
                        }
                        break;
                    default:
                        List<string> fieldNames = RestAPIWrapper.GetCharacterFields(moduleName);

                        if (fieldNames.Contains("first_name") && fieldNames.Contains("last_name"))
                        {
                            /* This is Ian's suggestion */
                            queryText = ConstructQueryTextWithFirstAndLastNames(escapedSearchText, firstTerm, lastTerm);
                            foreach (string fieldName in new string[] { "first_name", "last_name" })
                            {
                                fieldsToSeek.Add(fieldName);
                                fieldsToOrderBy.Add(new Tuple<string, string>(this.moduleName, fieldName));
                            }
                        }
                        else
                        {
                            queryText = ConstructQueryTextForUnknownModule(moduleName, escapedSearchText);

                            foreach (string fieldName in new string[] { "name", "description" })
                            {
                                if (fieldNames.Contains(fieldName))
                                {
                                    fieldsToSeek.Add(fieldName);
                                    fieldsToOrderBy.Add(new Tuple<string, string>(this.moduleName, fieldName));
                                }
                            }
                        }
                        break;
                }

                if (string.IsNullOrEmpty(queryText))
                {
                    throw new CouldNotConstructQueryException(moduleName, searchText);
                }

                foreach (var replacement in this.replacements)
                {
                    queryText = replacement.Item1.Replace(queryText, replacement.Item2);
                }

                return queryText;
            }

            /// <summary>
            /// Construct an array of links to traverse to the target module.
            /// </summary>
            /// <returns>The constructed array.</returns>
            private object ConstructLinkNamesToFieldsArray()
            {
                object result;
                try
                {
                    ICollection<LinkSpec> chain = this.moduleChains[moduleName];
                    ICollection<object> links = new List<object>();

                    foreach (var link in chain)
                    {
                        links.Add(new { @name = link.LinkName, @value = link.FieldNames });
                    }

                    result = links.ToArray();
                }
                catch (KeyNotFoundException)
                {
                    result = null;
                }

                return result;
            }

            /// <summary>
            /// Execute me on the server.
            /// </summary>
            /// <returns>The results of executing me.</returns>
            internal EntryList Execute()
            {
                string queryText = this.QueryText;

                return RestAPIWrapper.GetEntryList(moduleName,
                    queryText,
                    Properties.Settings.Default.SyncMaxRecords,
                    this.OrderClause,
                    0,
                    false,
                    fieldsToSeek.ToArray(),
                    ConstructLinkNamesToFieldsArray());
            }

            /// <summary>
            /// Replace this pattern with this substitution in my query text.
            /// </summary>
            /// <remarks>Note that while you can add replacements, you cannot 
            /// remove them.</remarks>
            /// <param name="pattern">The pattern to replace.</param>
            /// <param name="substitution">The substitution to replace it with.</param>
            /// <returns>Myself, to allow chaining.</returns>
            public SearchQuery Replace(string pattern, string substitution)
            {
                return this.Replace(new Regex(pattern), substitution);
            }

            /// <summary>
            /// Replace this pattern with this substitution in my query text.
            /// </summary>
            /// <remarks>Note that while you can add replacements, you cannot 
            /// remove them.</remarks>
            /// <param name="pattern">The pattern to replace.</param>
            /// <param name="substitution">The substitution to replace it with.</param>
            /// <returns>Myself, to allow chaining.</returns>
            public SearchQuery Replace(Regex attern, string substitution)
            {
                this.replacements.Add(new Tuple<Regex, string>(attern, substitution));
                return this;
            }
        }
    }
}
