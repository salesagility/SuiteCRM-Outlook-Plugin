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
using SuiteCRMAddIn.Daemon;
using SuiteCRMAddIn.Extensions;
using SuiteCRMAddIn.Helpers;
using SuiteCRMClient.RESTObjects;

#endregion

namespace SuiteCRMAddIn.Dialogs
{
    public partial class ManualSyncContactForm : Form
    {
        private Dictionary<string, EntryValue> searchResults = new Dictionary<string, EntryValue>();

        public ManualSyncContactForm(string searchString)
        {
            InitializeComponent();
            searchText.Text = searchString;
            ClearAndSearch(searchString);
        }

        private void ClearAndSearch(string target)
        {
            searchResults = new Dictionary<string, EntryValue>();

            resultsTree.Nodes.Clear();
            resultsTree.Nodes.Add("Create", "Create a new Contact");

            if (!string.IsNullOrWhiteSpace(target))
            {
                var contactsNode = resultsTree.Nodes.Add("Contacts", "Contacts");

                SearchAddChildren(target, contactsNode);
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
            if (resultsTree.Nodes["create"].Checked)
            {
                foreach (var contactItem in Globals.ThisAddIn.SelectedContacts)
                {
                    var state =
                        (ContactSyncState) SyncStateManager.Instance.GetOrCreateSyncState(contactItem);
                    if (!state.ExistedInCrm)
                    {
                        contactItem.SetManualOverride();
                        state.SetPending();
                        DaemonWorker.Instance.AddTask(new TransmitUpdateAction<ContactItem, ContactSyncState>(
                            Globals.ThisAddIn.ContactsSynchroniser,
                            state));

                        Close();
                    }
                    else
                    {
                        MessageBox.Show($"Contact {contactItem.FullName} already exists in CRM", "Contact Exists",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            else
            {
                foreach (var contactItem in Globals.ThisAddIn.SelectedContacts)
                {
                    var state =
                        (ContactSyncState) SyncStateManager.Instance.GetOrCreateSyncState(contactItem);
                    if (resultsTree.Nodes[state.CrmEntryId.ToString()].Checked)
                    {
                        contactItem.SetManualOverride();
                        state.SetPending();
                        DaemonWorker.Instance.AddTask(new TransmitUpdateAction<ContactItem, ContactSyncState>(
                            Globals.ThisAddIn.ContactsSynchroniser,
                            state));
                    }
                }

                Close();
            }
        }

        private void cancelButton_click(object sender, EventArgs e)
        {
            Close();
        }

        private void resultsTree_ItemClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node == resultsTree.Nodes["Create"])
            {
                if (e.Node.Checked)
                    foreach (var node in resultsTree.Nodes["Contacts"].GetAllNodes())
                        node.Checked = false;
            }
            else
            {
                if (e.Node.Checked)
                    resultsTree.Nodes["Create"].Checked = false;
            }

            saveButton.Enabled = resultsTree.GetAllNodes().Any(x => x.Checked);
        }
    }
}