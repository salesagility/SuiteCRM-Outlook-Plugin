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
using System.Drawing;
using System.Windows.Forms;
using SuiteCRMClient;
using SuiteCRMClient.Logging;
using SuiteCRMClient.RESTObjects;
using ListViewEx;
using System.Collections.Specialized;

namespace SuiteCRMAddIn
{
    public partial class frmCustomModules : Form
    {
        private clsSettings settings = Globals.ThisAddIn.settings;
        private List<string> IgnoreModules = new List<string>();
        private TextBox txtDisplay;

        public frmCustomModules()
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
            this.IgnoreModules.Add("Contacts");
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
            this.IgnoreModules.Add("Calls");
            this.IgnoreModules.Add("Tasks");
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
            this.IgnoreModules.Add("Meetings");
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

        private void listViewAvailableModules_SubItemClicked(object sender, SubItemEventArgs e)
        {
            try
            {
                if (e.SubItem == 1)
                {
                    this.lstViewAvailableModules.StartEditing(this.txtDisplay, e.Item, e.SubItem);
                }
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "listViewAvailableModules_SubItemClicked General Exception:\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------\n";
                Globals.ThisAddIn.Log.Warn(strLog);
            }
        }

        private void frmCustomModules_Load(object sender, EventArgs e)
        {
            try
            {
                clsSuiteCRMHelper.EnsureLoggedIn(Globals.ThisAddIn.SuiteCRMUserSession);

                if (Globals.ThisAddIn.SuiteCRMUserSession.id == "")
                {
                    MessageBox.Show("Please enter SuiteCRM details in General tab and try again", "Invalid Authentication");
                    base.Close();
                    return;
                }
                   eModuleList modules = clsSuiteCRMHelper.GetModules();
                       this.lstViewAvailableModules.SubItemClicked += new SubItemEventHandler(this.listViewAvailableModules_SubItemClicked);
                       if (this.settings.CustomModules != null)
                       {
                           StringEnumerator enumerator = this.settings.CustomModules.GetEnumerator();
                           while (enumerator.MoveNext())
                           {
                               string[] strArray = enumerator.Current.Split(new char[] { '|' });
                               ListViewItem item = new ListViewItem
                               {
                                   Text = strArray[0],
                                   Tag = strArray[1],
                                   Checked = true
                               };
                               item.SubItems.Add(strArray[1]);
                               if (strArray[0] != "None" || strArray[1] != "None")
                                   this.lstViewAvailableModules.Items.Add(item);
                           }
                       }
                       foreach (module_data objModuleData in modules.modules1)
                       {
                           string str2 = objModuleData.module_key;
                           bool flag = false;
                           if (!this.IgnoreModules.Contains(str2))
                           {
                               ListViewItem item2 = new ListViewItem
                               {
                                   Text = str2,
                                   Tag = str2
                               };
                               item2.SubItems.Add(string.Empty);
                               foreach (ListViewItem item3 in this.lstViewAvailableModules.Items)
                               {
                                   if (item3.Text == str2)
                                   {
                                       flag = true;
                                   }
                               }
                               if (!flag)
                               {
                                   this.lstViewAvailableModules.Items.Add(item2);
                               }
                           }
                       }
            }
            catch (Exception ex)
            {
                base.Close();
                MessageBox.Show("Please check the Internet connection", "Network Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "frmCustomModules_Load General Exception\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "Data:" + ex.Data.ToString() + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------\n";
                Globals.ThisAddIn.Log.Warn(strLog);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                bool flag = true;
                this.settings.CustomModules.Clear();
                foreach (ListViewItem item in this.lstViewAvailableModules.CheckedItems)
                {
                    if (item.SubItems[1].Text != string.Empty)
                    {
                        this.settings.CustomModules.Add(item.Text + "|" + item.SubItems[1].Text);
                    }
                    else
                    {
                        flag = false;
                    }
                }
                if (flag)
                {
                    this.settings.Save();
                    this.settings.Reload();
                    base.Close();
                }
                else
                {
                    MessageBox.Show("You have not entered a label for all of the selected modules.");
                }
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "buttonSaveClose_Click General Exception:\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------\n";
                Globals.ThisAddIn.Log.Warn(strLog);
                // Swallow exception(!)
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            base.Close();
        }
    }
}
