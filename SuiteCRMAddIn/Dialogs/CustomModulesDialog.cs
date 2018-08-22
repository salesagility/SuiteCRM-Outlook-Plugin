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
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using SuiteCRMClient;
using SuiteCRMClient.Logging;
using SuiteCRMClient.RESTObjects;
using SuiteCRMAddIn.BusinessLogic;

namespace SuiteCRMAddIn.Dialogs
{
    public partial class CustomModulesDialog : Form
    {
        private List<string> IgnoreModules = new List<string>();
        private TextBox txtDisplay;

        public CustomModulesDialog()
        {
            InitializeComponent();
            this.txtDisplay = new TextBox();
            this.txtDisplay.Location = new Point(0xda, 0x1b);
            this.txtDisplay.Name = "textBoxDisplay";
            this.txtDisplay.Size = new Size(100, 0x15);
            this.txtDisplay.TabIndex = 11;
            this.txtDisplay.Visible = false;
            base.Controls.Add(this.txtDisplay);
            this.IgnoreModules.Add("iFrames");
            this.IgnoreModules.Add(ContactSynchroniser.CrmModule);
            this.IgnoreModules.Add("Accounts");
            this.IgnoreModules.Add("Projects");
            this.IgnoreModules.Add("Bugs");
            this.IgnoreModules.Add("Opportunities");
            this.IgnoreModules.Add("Cases");
            this.IgnoreModules.Add("Leads");
            this.IgnoreModules.Add("Trackers");
            this.IgnoreModules.Add("Project");
            this.IgnoreModules.Add("KBDocuments");
            this.IgnoreModules.Add("Trackers");
            this.IgnoreModules.Add("Emails");
            this.IgnoreModules.Add(CallsSynchroniser.CrmModule);
            this.IgnoreModules.Add(TaskSynchroniser.CrmModule);
            this.IgnoreModules.Add("UserPrefs");
            this.IgnoreModules.Add("Contracts");
            this.IgnoreModules.Add("Campaigns");
            this.IgnoreModules.Add("Documents");
            this.IgnoreModules.Add("Quotes");
            this.IgnoreModules.Add("Products");
            this.IgnoreModules.Add("Forecasts");
            this.IgnoreModules.Add("Reports");
            this.IgnoreModules.Add("Feeds");
            this.IgnoreModules.Add("Administration");
            this.IgnoreModules.Add("Currencies");
            this.IgnoreModules.Add("EditCustomFields");
            this.IgnoreModules.Add("Manufacturers");
            this.IgnoreModules.Add("ProductBundles");
            this.IgnoreModules.Add("ProductBundleNotes");
            this.IgnoreModules.Add("ProductCategories");
            this.IgnoreModules.Add("ProductTemplates");
            this.IgnoreModules.Add("ProductTypes");
            this.IgnoreModules.Add("Shippers");
            this.IgnoreModules.Add("TaxRates");
            this.IgnoreModules.Add("TeamNotices");
            this.IgnoreModules.Add("Teams");
            this.IgnoreModules.Add("TimePeriods");
            this.IgnoreModules.Add("ForecastOpportunities");
            this.IgnoreModules.Add("Quotas");
            this.IgnoreModules.Add("KBDocumentRevisions");
            this.IgnoreModules.Add("KBDocumentKBTags");
            this.IgnoreModules.Add("KBTags");
            this.IgnoreModules.Add("KBTags");
            this.IgnoreModules.Add("KBContents");
            this.IgnoreModules.Add("Users");
            this.IgnoreModules.Add("Versions");
            this.IgnoreModules.Add("Roles");
            this.IgnoreModules.Add("EmailMarketing");
            this.IgnoreModules.Add("TeamMemberships");
            this.IgnoreModules.Add("TeamSets");
            this.IgnoreModules.Add("MergeRecords");
            this.IgnoreModules.Add("EmailAddresses");
            this.IgnoreModules.Add("Schedulers");
            this.IgnoreModules.Add("EmailTemplates");
            this.IgnoreModules.Add("CampaignTrackers");
            this.IgnoreModules.Add("CampaignLog");
            this.IgnoreModules.Add("EmailMan");
            this.IgnoreModules.Add("Prospects");
            this.IgnoreModules.Add("ProspectLists");
            this.IgnoreModules.Add("InboundEmail");
            this.IgnoreModules.Add("ACLActions");
            this.IgnoreModules.Add("ACLRoles");
            this.IgnoreModules.Add("DocumentRevisions");
            this.IgnoreModules.Add("ContractTypes");
            this.IgnoreModules.Add("ForecastSchedule");
            this.IgnoreModules.Add("Worksheet");
            this.IgnoreModules.Add("ACLFields");
            this.IgnoreModules.Add("ProjectResources");
            this.IgnoreModules.Add("Holidays");
            this.IgnoreModules.Add("ProjectTask");
            this.IgnoreModules.Add("WorkFlow");
            this.IgnoreModules.Add("WorkFlowTriggerShells");
            this.IgnoreModules.Add("WorkFlowAlertShells");
            this.IgnoreModules.Add("WorkFlowAlerts");
            this.IgnoreModules.Add("WorkFlowActionShells");
            this.IgnoreModules.Add("WorkFlowActions");
            this.IgnoreModules.Add("Expressions");
            this.IgnoreModules.Add("UserPreferences");
            this.IgnoreModules.Add("SavedSearch");
            this.IgnoreModules.Add("SugarFeed");
            this.IgnoreModules.Add("SugarFavorites");
            this.IgnoreModules.Add(MeetingsSynchroniser.CrmModule);
            this.IgnoreModules.Add("Notes");
            this.IgnoreModules.Add("TrackerPerfs");
            this.IgnoreModules.Add("TrackerQueries");
            this.IgnoreModules.Add("TrackerSessions");
            this.IgnoreModules.Add("Employees");
            this.IgnoreModules.Add("Groups");
            this.IgnoreModules.Add("Releases");
            this.IgnoreModules.Add("Home");
            this.IgnoreModules.Add("Calendar");
            this.IgnoreModules.Add("Activities");
            this.IgnoreModules.Add("CustomFields");
            this.IgnoreModules.Add("Connectors");
            this.IgnoreModules.Add("Dropdown");
            this.IgnoreModules.Add("Dynamic");
            this.IgnoreModules.Add("DynamicFields");
            this.IgnoreModules.Add("DynamicLayout");
            this.IgnoreModules.Add("Help");
            this.IgnoreModules.Add("Import");
            this.IgnoreModules.Add("MySettings");
            this.IgnoreModules.Add("FieldsMetaData");
            this.IgnoreModules.Add("UpgradeWizard");
            this.IgnoreModules.Add("Sync");
            this.IgnoreModules.Add("LabelEditor");
            this.IgnoreModules.Add("OptimisticLock");
            this.IgnoreModules.Add("Audit");
            this.IgnoreModules.Add("MailMerge");
            this.IgnoreModules.Add("Schedulers_jobs");
            this.IgnoreModules.Add("ACL");
            this.IgnoreModules.Add("Configurator");
            this.IgnoreModules.Add("Studio");
            this.IgnoreModules.Add("LoginAudit");
            this.IgnoreModules.Add("Search");
            this.IgnoreModules.Add("Dashboard");
            this.IgnoreModules.Add("EmailText");
            this.IgnoreModules.Add("Notifications");
            this.IgnoreModules.Add("EAPM");
            this.IgnoreModules.Add("OAuthKeys");
            this.IgnoreModules.Add("OAuthTokens");
            this.IgnoreModules.Add("Dashboard");
        }

        private ILogger Log => Globals.ThisAddIn.Log;

        private void frmCustomModules_Load(object sender, EventArgs e)
        {
            using (WaitCursor.For(this))
            {
                try
                {
                    RestAPIWrapper.EnsureLoggedIn();

                    if (Globals.ThisAddIn.SuiteCRMUserSession.NotLoggedIn)
                    {
                        MessageBox.Show("Please enter SuiteCRM details in General tab and try again", "Invalid Authentication");
                        base.Close();
                        return;
                    }

                    PopulateCustomModulesListView(this.lstViewAvailableModules, this.IgnoreModules);
                }
                catch (Exception ex)
                {
                    Log.Warn("frmCustomModules_Load error", ex);
                    base.Close();
                    MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// Populate this list view with custom modules, marking those saved in my settings as selected.
        /// </summary>
        /// <param name="view">The view to populate.</param>
        /// <param name="toIgnore">Keys of modules to ignore.</param>
        protected void PopulateCustomModulesListView(ListView view, List<string> toIgnore)
        {
            foreach (AvailableModule module in 
                RestAPIWrapper.GetModulesHavingEmailRelationships()
                .OrderBy(i => i.module_key))
            {
                if (!toIgnore.Contains(module.module_key))
                {
                    view.Items.Add(new ListViewItem
                        {
                            Checked = IsSelectedCustomModule(module),
                            Text = module.module_key,
                            Tag = module.module_key,
                            SubItems = { string.IsNullOrWhiteSpace(module.module_label) ?
                                            module.module_key :
                                            module.module_label}
                        });
                }
            }
        }

        /// <summary>
        /// Is this module a currently selected custom module?
        /// </summary>
        /// <param name="module">The module.</param>
        /// <returns>True if this module is a currently selected custom module.</returns>
        private bool IsSelectedCustomModule(AvailableModule module)
        {
            return Properties.Settings.Default.CustomModules != null &&
                Properties.Settings.Default.CustomModules.Where(i => i.StartsWith($"{module.module_key}|")).Count() > 0;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                bool flag = true;
                Properties.Settings.Default.CustomModules = new List<string>();
                foreach (ListViewItem item in this.lstViewAvailableModules.CheckedItems)
                {
                    if (item.SubItems[1].Text != string.Empty)
                    {
                        Properties.Settings.Default.CustomModules.Add(item.Text + "|" + item.SubItems[1].Text);
                    }
                    else
                    {
                        flag = false;
                    }
                }
                if (flag)
                {
                    Properties.Settings.Default.Save();
                    Properties.Settings.Default.Reload();
                    base.Close();
                }
                else
                {
                    MessageBox.Show("You have not entered a label for all of the selected modules.");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle("Failure while trying to save selected custom modules", ex);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            base.Close();
        }
    }
}
