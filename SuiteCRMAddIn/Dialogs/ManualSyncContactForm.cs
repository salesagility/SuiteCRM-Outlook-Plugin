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

#region

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SuiteCRMAddIn.BusinessLogic;
using SuiteCRMAddIn.Extensions;
using SuiteCRMAddIn.Helpers;
using SuiteCRMAddIn.Properties;
using SuiteCRMClient.RESTObjects;

#endregion

namespace SuiteCRMAddIn.Dialogs
{
    public partial class ManualSyncContactForm : Form
    {
        /// <summary>
        ///     The key for the create node.
        /// </summary>
        private static readonly string CreateNodeKey = "Create";

        /// <summary>
        ///     The key for the contacts node.
        /// </summary>
        private static readonly string ContactsNodeKey = "Contacts";

        private Dictionary<string, EntryValue> searchResults = new Dictionary<string, EntryValue>();

        public ManualSyncContactForm(string searchString)
        {
            InitializeComponent();
            searchText.Text = searchString;
        }

        private void ClearAndSearch(string target)
        {
            using (WaitCursor.For(this, true))
            {
                searchResults = new Dictionary<string, EntryValue>();

                resultsTree.Nodes.Clear();
                resultsTree.Nodes.Add(CreateNodeKey, "Create a new Contact");

                if (!string.IsNullOrWhiteSpace(target))
                {
                    var contactsNode = resultsTree.Nodes.Add(ContactsNodeKey, "Contacts");

                    SearchAddChildren(target, contactsNode);

                    if (contactsNode.Nodes.Count == 0)
                    {
                        resultsTree.Nodes.Remove(contactsNode);
                        resultsTree.Nodes[CreateNodeKey].Checked = true;
                        saveButton.Enabled = true;
                    }
                }
            }
        }

        private void SearchAddChildren(string target, TreeNode contactsNode)
        {
            var tokens = target.Split(" ;:,".ToCharArray());

            foreach (var token in tokens.Where(x => !string.IsNullOrEmpty(x)))
            foreach (var crmContact in SearchHelper.SearchContacts(token))
                searchResults[crmContact.id] = crmContact;

            foreach (var result in searchResults.Values.OrderBy(
                x => $"{x.GetValueAsString("last_name")} {x.GetValueAsString("first_name")}"))
                contactsNode.Nodes.Add(result.id, CanonicalString(result));

            contactsNode.Expand();
        }

        private static string CanonicalString(EntryValue result)
        {
            return
                $"{result.GetValueAsString("first_name")} {result.GetValueAsString("last_name")} ({result.GetValueAsString("email1")})";
        }

        private void searchButton_click(object sender, EventArgs e)
        {
            ClearAndSearch(searchText.Text);
        }

        private void saveButton_click(object sender, EventArgs e)
        {
            var shouldClose = true;

            foreach (var contactItem in Globals.ThisAddIn.SelectedContacts)
            {
                var state =
                    (ContactSyncState) SyncStateManager.Instance.GetOrCreateSyncState(contactItem);
                var proceed = true;
                var crmId = contactItem.GetCrmId().ToString();
                var synchroniser = Globals.ThisAddIn.ContactsSynchroniser;

                if (contactItem.Sensitivity == OlSensitivity.olPrivate)
                    if (MessageBox.Show($"Contact {contactItem.FullName} is marked 'private'. Are you sure?",
                            "Private: are you sure?", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) ==
                        DialogResult.Cancel)
                    {
                        proceed = false;
                        shouldClose = false;
                    }
                if (proceed)
                    if (resultsTree.Nodes["create"].Checked)
                    {
                        if (!state.ExistedInCrm)
                        {
                            contactItem.SetManualOverride();
                        }
                        else
                        {
                            MessageBox.Show($"Contact {contactItem.FullName} already exists in CRM", "Contact Exists",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            shouldClose = false;
                        }
                    }
                    else if (searchResults.ContainsKey(crmId))
                    {
                        contactItem.SetManualOverride();
                    }
                    else if (string.IsNullOrEmpty(crmId))
                    {
                        var p = contactItem.UserProperties[SyncStateManager.CrmIdPropertyName] ??
                                contactItem.UserProperties.Add(SyncStateManager.CrmIdPropertyName,
                                    OlUserPropertyType.olText);
                        try
                        {
                            p.Value = resultsTree.Nodes[ContactsNodeKey].Nodes[0].Name;
                            state.CrmEntryId = p.Value;
                            contactItem.Save();
                        }
                        finally
                        {
                            contactItem.SetManualOverride();
                        }
                    }
            }

            if (shouldClose) Close();
        }


        private void cancelButton_click(object sender, EventArgs e)
        {
            Close();
        }

        private void resultsTree_ItemClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            var contactsNode = resultsTree.Nodes[ContactsNodeKey];
            var createNode = resultsTree.Nodes[CreateNodeKey];

            if (e.Node == createNode)
            {
                if (e.Node.Checked && contactsNode != null)
                    foreach (var node in contactsNode.GetAllNodes())
                        node.Checked = false;
            }
            else if (e.Node == contactsNode)
            {
                if (e.Node.Checked)
                    foreach (var node in contactsNode.GetAllNodes())
                        node.Checked = true;
            }
            else
            {
                if (e.Node.Checked)
                    createNode.Checked = false;
                else
                    contactsNode.Checked = false;
            }

            saveButton.Enabled = resultsTree.GetAllNodes().Any(x => x.Checked);
        }

        private void manualSyncContactsForm_Load(object sender, EventArgs e)
        {
            if (Settings.Default.AutomaticSearch)
                BeginInvoke((MethodInvoker) delegate { ClearAndSearch(searchText.Text); });
        }

        private void seachText_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !string.IsNullOrWhiteSpace(searchText.Text))
                ClearAndSearch(searchText.Text);
        }
    }
}