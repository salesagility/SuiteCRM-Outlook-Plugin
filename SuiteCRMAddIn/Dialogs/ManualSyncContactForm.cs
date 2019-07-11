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
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SuiteCRMAddIn.BusinessLogic;
using SuiteCRMAddIn.Extensions;
using SuiteCRMAddIn.Helpers;
using SuiteCRMAddIn.Properties;
using SuiteCRMClient.RESTObjects;
using System.Text;

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

        private bool dontClose = false;

        /// <summary>
        /// The contact we're going to operate on.
        /// </summary>
        ContactItem contactItem = Globals.ThisAddIn.SelectedContacts.First();

        public ManualSyncContactForm(string searchString)
        {
            if (contactItem != null)
            {
                InitializeComponent();
                this.Text = $"Manually sync {contactItem.FullName}";
                searchText.Text = searchString;
            }
            else
            {
                throw new System.Exception("No contact selected in ManualSyncContactForm");
            }
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
            {
                TreeNode node = contactsNode.Nodes.Add(result.id, CanonicalString(result));
                var contactItem = Globals.ThisAddIn.SelectedContacts.First();

                if (IsProbablySameItem(result, contactItem))
                {
                    node.BackColor = ColorTranslator.FromHtml("#a9ea56");
                }
                else if (IsPreviouslySyncedItem(result))
                {
                    node.BackColor = ColorTranslator.FromHtml("#ea6556");
                }
                else if (SyncStateManager.Instance.GetExistingSyncState(result) != null)
                {
                    node.BackColor = ColorTranslator.FromHtml("#ea6556");
                }

                contactsNode.Expand();
            }
        }

        private bool IsPreviouslySyncedItem(string crmId)
        {
            return !string.IsNullOrEmpty(crmId) &&
                   searchResults.ContainsKey(crmId) &&
                   IsPreviouslySyncedItem(searchResults[crmId]);
        }

        private bool IsPreviouslySyncedItem(EntryValue result)
        {
            return !string.IsNullOrEmpty(result.GetValueAsString("outlook_id")) ||
                   !string.IsNullOrEmpty(result.GetValueAsString("sync_contact")) ||
                   SyncStateManager.Instance.GetExistingSyncState(result) != null;
        }

        private bool IsProbablySameItem(EntryValue result, ContactItem contactItem)
        {
            string crmIdStr = contactItem.GetCrmId().ToString();
            return result != null &&
                (result.id.Equals(crmIdStr) ||
                   result.GetValueAsString("outlook_id").Equals(contactItem.EntryID));
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
            var crmId = contactItem.GetCrmId().ToString();
            string selectedId = resultsTree.GetAllNodes().FirstOrDefault(x => x.Checked)?.Name;
            EntryValue selectedItem = searchResults.ContainsKey(selectedId) ? searchResults[selectedId] : null;
            List<string> problems = new List<string>();

            if (contactItem.Sensitivity == OlSensitivity.olPrivate)
            {
                problems.Add($"Contact {contactItem.FullName} is marked 'private'. Are you sure?");
            }

            if (resultsTree.Nodes["create"].Checked && IsPreviouslySyncedItem(crmId))
            {
                problems.Add($"A record for contact {contactItem.FullName} already exists in CRM. Are you sure you want to create a new record?");
            }
            if (selectedItem != null && 
                     !IsProbablySameItem(selectedItem, contactItem))
            {
                problems.Add($"The record for {selectedItem.GetValueAsString("first_name")} {selectedItem.GetValueAsString("last_name")} will be overwritten with the details of {contactItem.FullName}.");
            }
            if (IsPreviouslySyncedItem(crmId) && selectedItem != null)
            {
                problems.Add($"Contact {selectedItem.GetValueAsString("first_name")} {selectedItem.GetValueAsString("last_name")} has previously been synced and will be overwritten.");
            }

            if (resultsTree.Nodes["create"].Checked &&
                     IsPreviouslySyncedItem(crmId) )
            {
                problems.Add($"Contact {contactItem.FullName} has previously been synced. Are you sure you want to create another copy?");
            }

            if (problems.Count == 0 || MessageBox.Show(
                    string.Join("\n", problems.Select(p => $"• {p}\n").ToArray()),
                    "Problems found: are you sure?",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Warning) ==
                DialogResult.OK)
            {
                if (resultsTree.Nodes["create"].Checked)
                {
                    contactItem.ClearCrmId();
                    contactItem.SetManualOverride();
                }
                else
                {
                    try
                    {
                        contactItem.ChangeCrmId(resultsTree.GetAllNodes().FirstOrDefault(x => x.Checked).Name);
                    }
                    finally
                    {
                        contactItem.SetManualOverride();
                    }
                }
            }
            else
            {
                dontClose = true;
            }
        }

        private void cancelButton_click(object sender, EventArgs e)
        {
            Close();
        }

        private void resultsTree_ItemClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            var contactsNode = resultsTree.Nodes[ContactsNodeKey];
            var createNode = resultsTree.Nodes[CreateNodeKey];

            if (e.Node == contactsNode)
            {
                e.Node.Checked = false;
                // You can't check the 'Contacts' node.
            }
            else
            {
                foreach (var node in resultsTree.GetAllNodes().Where( n => n != e.Node))
                    node.Checked = false;
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

        private void FormClosingEvent(object sender, FormClosingEventArgs e)
        {
            e.Cancel = dontClose;
            dontClose = false;
        }
    }
}
