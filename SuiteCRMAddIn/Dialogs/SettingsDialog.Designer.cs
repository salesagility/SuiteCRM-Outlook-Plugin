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
    partial class SettingsDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsDialog));
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dtpAutoArchiveFrom = new System.Windows.Forms.DateTimePicker();
            this.label8 = new System.Windows.Forms.Label();
            this.EmailArchiveAccountTabs = new System.Windows.Forms.TabControl();
            this.autoArchiveAccountsPage = new System.Windows.Forms.TabPage();
            this.label7 = new System.Windows.Forms.Label();
            this.txtAutoSync = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.advancedArchiveSettingsButton = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkShowConfirmationMessageArchive = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnSelect = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.txtSyncMaxRecords = new System.Windows.Forms.TextBox();
            this.cbShowCustomModules = new System.Windows.Forms.CheckBox();
            this.checkBoxShowRightClick = new System.Windows.Forms.CheckBox();
            this.checkBoxAutomaticSearch = new System.Windows.Forms.CheckBox();
            this.cbEmailAttachments = new System.Windows.Forms.CheckBox();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.licenceGroup = new System.Windows.Forms.GroupBox();
            this.licenceText = new System.Windows.Forms.TextBox();
            this.licenceLabel = new System.Windows.Forms.Label();
            this.gbLDAPAuthentication = new System.Windows.Forms.GroupBox();
            this.labelKey = new System.Windows.Forms.Label();
            this.txtLDAPAuthenticationKey = new System.Windows.Forms.TextBox();
            this.chkEnableLDAPAuthentication = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnTestLogin = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUsername = new System.Windows.Forms.TextBox();
            this.txtURL = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.synchronisationTab = new System.Windows.Forms.TabPage();
            this.syncTasksMenu = new System.Windows.Forms.ComboBox();
            this.syncTasksLabel = new System.Windows.Forms.Label();
            this.syncMeetingsMenu = new System.Windows.Forms.ComboBox();
            this.syncMeetingsLabel = new System.Windows.Forms.Label();
            this.syncCallsMenu = new System.Windows.Forms.ComboBox();
            this.syncCallsLabel = new System.Windows.Forms.Label();
            this.syncContactsMenu = new System.Windows.Forms.ComboBox();
            this.syncContactsLabel = new System.Windows.Forms.Label();
            this.syncLabel = new System.Windows.Forms.Label();
            this.InformationTabPage = new System.Windows.Forms.TabPage();
            this.dotNetVersionLabel = new System.Windows.Forms.Label();
            this.AddInVersionLabel = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.crmIdValidationSelector = new System.Windows.Forms.ComboBox();
            this.policyBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.crmIdValidationLabel = new System.Windows.Forms.Label();
            this.showErrorsSelector = new System.Windows.Forms.ComboBox();
            this.popupWhenBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.showErrorsLabel = new System.Windows.Forms.Label();
            this.logLevelHelp = new System.Windows.Forms.Label();
            this.logLevelLabel = new System.Windows.Forms.Label();
            this.logLevelSelector = new System.Windows.Forms.ComboBox();
            this.LinkToLogFileDir = new System.Windows.Forms.LinkLabel();
            this.label11 = new System.Windows.Forms.Label();
            this.AddInTitleLabel = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.startupDeferralGroup = new System.Windows.Forms.GroupBox();
            this.startupDeferralLabel = new System.Windows.Forms.Label();
            this.startupDeferralInput = new System.Windows.Forms.NumericUpDown();
            this.startupDeferralHelp = new System.Windows.Forms.Label();
            this.tabPage3.SuspendLayout();
            this.EmailArchiveAccountTabs.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.licenceGroup.SuspendLayout();
            this.gbLDAPAuthentication.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.synchronisationTab.SuspendLayout();
            this.InformationTabPage.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.policyBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.popupWhenBindingSource)).BeginInit();
            this.startupDeferralGroup.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.startupDeferralInput)).BeginInit();
            this.SuspendLayout();
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dtpAutoArchiveFrom);
            this.tabPage3.Controls.Add(this.label8);
            this.tabPage3.Controls.Add(this.EmailArchiveAccountTabs);
            this.tabPage3.Controls.Add(this.label7);
            this.tabPage3.Controls.Add(this.txtAutoSync);
            this.tabPage3.Location = new System.Drawing.Point(23, 4);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(360, 374);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Auto Archive";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dtpAutoArchiveFrom
            // 
            this.dtpAutoArchiveFrom.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpAutoArchiveFrom.Location = new System.Drawing.Point(199, 340);
            this.dtpAutoArchiveFrom.Name = "dtpAutoArchiveFrom";
            this.dtpAutoArchiveFrom.Size = new System.Drawing.Size(94, 20);
            this.dtpAutoArchiveFrom.TabIndex = 24;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(12, 346);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(131, 13);
            this.label8.TabIndex = 0;
            this.label8.Text = "Automatically Archive from";
            // 
            // EmailArchiveAccountTabs
            // 
            this.EmailArchiveAccountTabs.Controls.Add(this.autoArchiveAccountsPage);
            this.EmailArchiveAccountTabs.Location = new System.Drawing.Point(11, 3);
            this.EmailArchiveAccountTabs.Name = "EmailArchiveAccountTabs";
            this.EmailArchiveAccountTabs.SelectedIndex = 0;
            this.EmailArchiveAccountTabs.Size = new System.Drawing.Size(340, 246);
            this.EmailArchiveAccountTabs.TabIndex = 17;
            // 
            // autoArchiveAccountsPage
            // 
            this.autoArchiveAccountsPage.Location = new System.Drawing.Point(4, 22);
            this.autoArchiveAccountsPage.Name = "autoArchiveAccountsPage";
            this.autoArchiveAccountsPage.Padding = new System.Windows.Forms.Padding(3);
            this.autoArchiveAccountsPage.Size = new System.Drawing.Size(332, 220);
            this.autoArchiveAccountsPage.TabIndex = 0;
            this.autoArchiveAccountsPage.Text = "Account#1";
            this.autoArchiveAccountsPage.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(8, 266);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(267, 13);
            this.label7.TabIndex = 14;
            this.label7.Text = "Exclude Messages From/To The Following Addresses :";
            // 
            // txtAutoSync
            // 
            this.txtAutoSync.AcceptsReturn = true;
            this.txtAutoSync.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtAutoSync.Location = new System.Drawing.Point(15, 282);
            this.txtAutoSync.Multiline = true;
            this.txtAutoSync.Name = "txtAutoSync";
            this.txtAutoSync.Size = new System.Drawing.Size(332, 52);
            this.txtAutoSync.TabIndex = 21;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.advancedArchiveSettingsButton);
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Location = new System.Drawing.Point(23, 4);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(360, 374);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Archive";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // advancedArchiveSettingsButton
            // 
            this.advancedArchiveSettingsButton.Location = new System.Drawing.Point(265, 306);
            this.advancedArchiveSettingsButton.Name = "advancedArchiveSettingsButton";
            this.advancedArchiveSettingsButton.Size = new System.Drawing.Size(75, 23);
            this.advancedArchiveSettingsButton.TabIndex = 1;
            this.advancedArchiveSettingsButton.Text = "Advanced";
            this.advancedArchiveSettingsButton.UseVisualStyleBackColor = true;
            this.advancedArchiveSettingsButton.Click += new System.EventHandler(this.advancedButton_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkShowConfirmationMessageArchive);
            this.groupBox2.Controls.Add(this.groupBox3);
            this.groupBox2.Controls.Add(this.checkBoxShowRightClick);
            this.groupBox2.Controls.Add(this.checkBoxAutomaticSearch);
            this.groupBox2.Controls.Add(this.cbEmailAttachments);
            this.groupBox2.Location = new System.Drawing.Point(7, 7);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(352, 282);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            // 
            // chkShowConfirmationMessageArchive
            // 
            this.chkShowConfirmationMessageArchive.AutoSize = true;
            this.chkShowConfirmationMessageArchive.Checked = true;
            this.chkShowConfirmationMessageArchive.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkShowConfirmationMessageArchive.Location = new System.Drawing.Point(26, 159);
            this.chkShowConfirmationMessageArchive.Name = "chkShowConfirmationMessageArchive";
            this.chkShowConfirmationMessageArchive.Size = new System.Drawing.Size(269, 17);
            this.chkShowConfirmationMessageArchive.TabIndex = 14;
            this.chkShowConfirmationMessageArchive.Text = "Show Confirmation Message on Successful Archive";
            this.chkShowConfirmationMessageArchive.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnSelect);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.txtSyncMaxRecords);
            this.groupBox3.Controls.Add(this.cbShowCustomModules);
            this.groupBox3.Location = new System.Drawing.Point(18, 63);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(325, 79);
            this.groupBox3.TabIndex = 10;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Search";
            // 
            // btnSelect
            // 
            this.btnSelect.Location = new System.Drawing.Point(258, 43);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(57, 23);
            this.btnSelect.TabIndex = 13;
            this.btnSelect.Text = "...";
            this.btnSelect.UseVisualStyleBackColor = true;
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 21);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(223, 13);
            this.label5.TabIndex = 5;
            this.label5.Text = "Maximum No. of Search Results Per Modules:";
            // 
            // txtSyncMaxRecords
            // 
            this.txtSyncMaxRecords.Location = new System.Drawing.Point(258, 18);
            this.txtSyncMaxRecords.Name = "txtSyncMaxRecords";
            this.txtSyncMaxRecords.Size = new System.Drawing.Size(57, 20);
            this.txtSyncMaxRecords.TabIndex = 11;
            // 
            // cbShowCustomModules
            // 
            this.cbShowCustomModules.AutoSize = true;
            this.cbShowCustomModules.Location = new System.Drawing.Point(8, 47);
            this.cbShowCustomModules.Name = "cbShowCustomModules";
            this.cbShowCustomModules.Size = new System.Drawing.Size(142, 17);
            this.cbShowCustomModules.TabIndex = 12;
            this.cbShowCustomModules.Text = "Include Custom Modules";
            this.cbShowCustomModules.UseVisualStyleBackColor = true;
            this.cbShowCustomModules.Click += new System.EventHandler(this.cbShowCustomModules_Click);
            // 
            // checkBoxShowRightClick
            // 
            this.checkBoxShowRightClick.AutoSize = true;
            this.checkBoxShowRightClick.Location = new System.Drawing.Point(26, 214);
            this.checkBoxShowRightClick.Name = "checkBoxShowRightClick";
            this.checkBoxShowRightClick.Size = new System.Drawing.Size(247, 17);
            this.checkBoxShowRightClick.TabIndex = 15;
            this.checkBoxShowRightClick.Text = "Show SuiteCRM Records in Right Click Menus";
            this.checkBoxShowRightClick.UseVisualStyleBackColor = true;
            // 
            // checkBoxAutomaticSearch
            // 
            this.checkBoxAutomaticSearch.AutoSize = true;
            this.checkBoxAutomaticSearch.Location = new System.Drawing.Point(26, 250);
            this.checkBoxAutomaticSearch.Name = "checkBoxAutomaticSearch";
            this.checkBoxAutomaticSearch.Size = new System.Drawing.Size(304, 17);
            this.checkBoxAutomaticSearch.TabIndex = 16;
            this.checkBoxAutomaticSearch.Text = "Automatically Search when the Archive Window is Opened";
            this.checkBoxAutomaticSearch.UseVisualStyleBackColor = true;
            // 
            // cbEmailAttachments
            // 
            this.cbEmailAttachments.AutoSize = true;
            this.cbEmailAttachments.Location = new System.Drawing.Point(26, 30);
            this.cbEmailAttachments.Name = "cbEmailAttachments";
            this.cbEmailAttachments.Size = new System.Drawing.Size(147, 17);
            this.cbEmailAttachments.TabIndex = 10;
            this.cbEmailAttachments.Text = "Archive Attachments Also";
            this.cbEmailAttachments.UseVisualStyleBackColor = true;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.licenceGroup);
            this.tabPage1.Controls.Add(this.gbLDAPAuthentication);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Location = new System.Drawing.Point(23, 4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(360, 374);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Credentials";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // licenceGroup
            // 
            this.licenceGroup.Controls.Add(this.licenceText);
            this.licenceGroup.Controls.Add(this.licenceLabel);
            this.licenceGroup.Location = new System.Drawing.Point(6, 258);
            this.licenceGroup.Name = "licenceGroup";
            this.licenceGroup.Size = new System.Drawing.Size(353, 52);
            this.licenceGroup.TabIndex = 4;
            this.licenceGroup.TabStop = false;
            // 
            // licenceText
            // 
            this.licenceText.Location = new System.Drawing.Point(78, 19);
            this.licenceText.Name = "licenceText";
            this.licenceText.Size = new System.Drawing.Size(207, 20);
            this.licenceText.TabIndex = 7;
            // 
            // licenceLabel
            // 
            this.licenceLabel.AutoSize = true;
            this.licenceLabel.Location = new System.Drawing.Point(6, 22);
            this.licenceLabel.Name = "licenceLabel";
            this.licenceLabel.Size = new System.Drawing.Size(66, 13);
            this.licenceLabel.TabIndex = 0;
            this.licenceLabel.Text = "Licence Key";
            // 
            // gbLDAPAuthentication
            // 
            this.gbLDAPAuthentication.Controls.Add(this.labelKey);
            this.gbLDAPAuthentication.Controls.Add(this.txtLDAPAuthenticationKey);
            this.gbLDAPAuthentication.Controls.Add(this.chkEnableLDAPAuthentication);
            this.gbLDAPAuthentication.Location = new System.Drawing.Point(6, 149);
            this.gbLDAPAuthentication.Name = "gbLDAPAuthentication";
            this.gbLDAPAuthentication.Size = new System.Drawing.Size(352, 103);
            this.gbLDAPAuthentication.TabIndex = 3;
            this.gbLDAPAuthentication.TabStop = false;
            // 
            // labelKey
            // 
            this.labelKey.AutoSize = true;
            this.labelKey.Enabled = false;
            this.labelKey.Location = new System.Drawing.Point(6, 44);
            this.labelKey.Name = "labelKey";
            this.labelKey.Size = new System.Drawing.Size(127, 13);
            this.labelKey.TabIndex = 2;
            this.labelKey.Text = "Password Encryption Key";
            // 
            // txtLDAPAuthenticationKey
            // 
            this.txtLDAPAuthenticationKey.Location = new System.Drawing.Point(78, 67);
            this.txtLDAPAuthenticationKey.Name = "txtLDAPAuthenticationKey";
            this.txtLDAPAuthenticationKey.Size = new System.Drawing.Size(207, 20);
            this.txtLDAPAuthenticationKey.TabIndex = 6;
            // 
            // chkEnableLDAPAuthentication
            // 
            this.chkEnableLDAPAuthentication.AutoSize = true;
            this.chkEnableLDAPAuthentication.Location = new System.Drawing.Point(9, 16);
            this.chkEnableLDAPAuthentication.Name = "chkEnableLDAPAuthentication";
            this.chkEnableLDAPAuthentication.Size = new System.Drawing.Size(161, 17);
            this.chkEnableLDAPAuthentication.TabIndex = 5;
            this.chkEnableLDAPAuthentication.Text = "Enable LDAP Authentication";
            this.chkEnableLDAPAuthentication.UseVisualStyleBackColor = true;
            this.chkEnableLDAPAuthentication.CheckedChanged += new System.EventHandler(this.chkEnableLDAPAuthentication_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnTestLogin);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txtPassword);
            this.groupBox1.Controls.Add(this.txtUsername);
            this.groupBox1.Controls.Add(this.txtURL);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(6, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(353, 137);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            // 
            // btnTestLogin
            // 
            this.btnTestLogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTestLogin.Location = new System.Drawing.Point(292, 103);
            this.btnTestLogin.Name = "btnTestLogin";
            this.btnTestLogin.Size = new System.Drawing.Size(42, 23);
            this.btnTestLogin.TabIndex = 4;
            this.btnTestLogin.Text = "&Test";
            this.btnTestLogin.UseVisualStyleBackColor = true;
            this.btnTestLogin.Click += new System.EventHandler(this.btnTestLogin_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(79, 184);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(207, 20);
            this.textBox2.TabIndex = 1;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(72, 40);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(169, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Ex : http://www.testcrm/suitecrm/";
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(79, 105);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(207, 20);
            this.txtPassword.TabIndex = 3;
            this.txtPassword.UseSystemPasswordChar = true;
            // 
            // txtUsername
            // 
            this.txtUsername.Location = new System.Drawing.Point(79, 71);
            this.txtUsername.Name = "txtUsername";
            this.txtUsername.Size = new System.Drawing.Size(207, 20);
            this.txtUsername.TabIndex = 2;
            // 
            // txtURL
            // 
            this.txtURL.Location = new System.Drawing.Point(79, 16);
            this.txtURL.Name = "txtURL";
            this.txtURL.Size = new System.Drawing.Size(207, 20);
            this.txtURL.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(7, 108);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Password";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Username";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "URL";
            // 
            // tabControl1
            // 
            this.tabControl1.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.synchronisationTab);
            this.tabControl1.Controls.Add(this.InformationTabPage);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Top;
            this.tabControl1.Location = new System.Drawing.Point(3, 3);
            this.tabControl1.Multiline = true;
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(387, 382);
            this.tabControl1.TabIndex = 43;
            // 
            // synchronisationTab
            // 
            this.synchronisationTab.Controls.Add(this.startupDeferralGroup);
            this.synchronisationTab.Controls.Add(this.syncTasksMenu);
            this.synchronisationTab.Controls.Add(this.syncTasksLabel);
            this.synchronisationTab.Controls.Add(this.syncMeetingsMenu);
            this.synchronisationTab.Controls.Add(this.syncMeetingsLabel);
            this.synchronisationTab.Controls.Add(this.syncCallsMenu);
            this.synchronisationTab.Controls.Add(this.syncCallsLabel);
            this.synchronisationTab.Controls.Add(this.syncContactsMenu);
            this.synchronisationTab.Controls.Add(this.syncContactsLabel);
            this.synchronisationTab.Controls.Add(this.syncLabel);
            this.synchronisationTab.Location = new System.Drawing.Point(23, 4);
            this.synchronisationTab.Name = "synchronisationTab";
            this.synchronisationTab.Padding = new System.Windows.Forms.Padding(3);
            this.synchronisationTab.Size = new System.Drawing.Size(360, 374);
            this.synchronisationTab.TabIndex = 4;
            this.synchronisationTab.Text = "Synchronisation";
            this.synchronisationTab.UseVisualStyleBackColor = true;
            // 
            // syncTasksMenu
            // 
            this.syncTasksMenu.FormattingEnabled = true;
            this.syncTasksMenu.Location = new System.Drawing.Point(126, 130);
            this.syncTasksMenu.Name = "syncTasksMenu";
            this.syncTasksMenu.Size = new System.Drawing.Size(228, 21);
            this.syncTasksMenu.TabIndex = 8;
            // 
            // syncTasksLabel
            // 
            this.syncTasksLabel.AutoSize = true;
            this.syncTasksLabel.Location = new System.Drawing.Point(38, 133);
            this.syncTasksLabel.Name = "syncTasksLabel";
            this.syncTasksLabel.Size = new System.Drawing.Size(39, 13);
            this.syncTasksLabel.TabIndex = 7;
            this.syncTasksLabel.Text = "Tasks:";
            // 
            // syncMeetingsMenu
            // 
            this.syncMeetingsMenu.FormattingEnabled = true;
            this.syncMeetingsMenu.Location = new System.Drawing.Point(126, 103);
            this.syncMeetingsMenu.Name = "syncMeetingsMenu";
            this.syncMeetingsMenu.Size = new System.Drawing.Size(228, 21);
            this.syncMeetingsMenu.TabIndex = 6;
            // 
            // syncMeetingsLabel
            // 
            this.syncMeetingsLabel.AutoSize = true;
            this.syncMeetingsLabel.Location = new System.Drawing.Point(38, 106);
            this.syncMeetingsLabel.Name = "syncMeetingsLabel";
            this.syncMeetingsLabel.Size = new System.Drawing.Size(53, 13);
            this.syncMeetingsLabel.TabIndex = 5;
            this.syncMeetingsLabel.Text = "Meetings:";
            // 
            // syncCallsMenu
            // 
            this.syncCallsMenu.FormattingEnabled = true;
            this.syncCallsMenu.Location = new System.Drawing.Point(126, 76);
            this.syncCallsMenu.Name = "syncCallsMenu";
            this.syncCallsMenu.Size = new System.Drawing.Size(228, 21);
            this.syncCallsMenu.TabIndex = 4;
            // 
            // syncCallsLabel
            // 
            this.syncCallsLabel.AutoSize = true;
            this.syncCallsLabel.Location = new System.Drawing.Point(38, 79);
            this.syncCallsLabel.Name = "syncCallsLabel";
            this.syncCallsLabel.Size = new System.Drawing.Size(32, 13);
            this.syncCallsLabel.TabIndex = 3;
            this.syncCallsLabel.Text = "Calls:";
            // 
            // syncContactsMenu
            // 
            this.syncContactsMenu.FormattingEnabled = true;
            this.syncContactsMenu.Location = new System.Drawing.Point(126, 49);
            this.syncContactsMenu.Name = "syncContactsMenu";
            this.syncContactsMenu.Size = new System.Drawing.Size(228, 21);
            this.syncContactsMenu.TabIndex = 2;
            // 
            // syncContactsLabel
            // 
            this.syncContactsLabel.AutoSize = true;
            this.syncContactsLabel.Location = new System.Drawing.Point(38, 52);
            this.syncContactsLabel.Name = "syncContactsLabel";
            this.syncContactsLabel.Size = new System.Drawing.Size(52, 13);
            this.syncContactsLabel.TabIndex = 1;
            this.syncContactsLabel.Text = "Contacts:";
            // 
            // syncLabel
            // 
            this.syncLabel.AutoSize = true;
            this.syncLabel.Location = new System.Drawing.Point(22, 24);
            this.syncLabel.Name = "syncLabel";
            this.syncLabel.Size = new System.Drawing.Size(68, 13);
            this.syncLabel.TabIndex = 0;
            this.syncLabel.Text = "Synchronise:";
            // 
            // InformationTabPage
            // 
            this.InformationTabPage.Controls.Add(this.dotNetVersionLabel);
            this.InformationTabPage.Controls.Add(this.AddInVersionLabel);
            this.InformationTabPage.Controls.Add(this.groupBox4);
            this.InformationTabPage.Controls.Add(this.AddInTitleLabel);
            this.InformationTabPage.Location = new System.Drawing.Point(23, 4);
            this.InformationTabPage.Name = "InformationTabPage";
            this.InformationTabPage.Size = new System.Drawing.Size(360, 374);
            this.InformationTabPage.TabIndex = 3;
            this.InformationTabPage.Text = "Information";
            this.InformationTabPage.UseVisualStyleBackColor = true;
            // 
            // dotNetVersionLabel
            // 
            this.dotNetVersionLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dotNetVersionLabel.Location = new System.Drawing.Point(13, 77);
            this.dotNetVersionLabel.Name = "dotNetVersionLabel";
            this.dotNetVersionLabel.Size = new System.Drawing.Size(332, 16);
            this.dotNetVersionLabel.TabIndex = 7;
            this.dotNetVersionLabel.Text = ".Net version 0.0.0";
            this.dotNetVersionLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // AddInVersionLabel
            // 
            this.AddInVersionLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AddInVersionLabel.Location = new System.Drawing.Point(13, 61);
            this.AddInVersionLabel.Name = "AddInVersionLabel";
            this.AddInVersionLabel.Size = new System.Drawing.Size(332, 16);
            this.AddInVersionLabel.TabIndex = 6;
            this.AddInVersionLabel.Text = "Version 0.0.0.0";
            this.AddInVersionLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.crmIdValidationSelector);
            this.groupBox4.Controls.Add(this.crmIdValidationLabel);
            this.groupBox4.Controls.Add(this.showErrorsSelector);
            this.groupBox4.Controls.Add(this.showErrorsLabel);
            this.groupBox4.Controls.Add(this.logLevelHelp);
            this.groupBox4.Controls.Add(this.logLevelLabel);
            this.groupBox4.Controls.Add(this.logLevelSelector);
            this.groupBox4.Controls.Add(this.LinkToLogFileDir);
            this.groupBox4.Controls.Add(this.label11);
            this.groupBox4.Location = new System.Drawing.Point(13, 144);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(332, 199);
            this.groupBox4.TabIndex = 5;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Logging";
            // 
            // crmIdValidationSelector
            // 
            this.crmIdValidationSelector.DataSource = this.policyBindingSource;
            this.crmIdValidationSelector.DisplayMember = "Value";
            this.crmIdValidationSelector.FormattingEnabled = true;
            this.crmIdValidationSelector.Location = new System.Drawing.Point(186, 160);
            this.crmIdValidationSelector.Name = "crmIdValidationSelector";
            this.crmIdValidationSelector.Size = new System.Drawing.Size(140, 21);
            this.crmIdValidationSelector.TabIndex = 38;
            // 
            // policyBindingSource
            // 
            this.policyBindingSource.DataSource = typeof(SuiteCRMAddIn.BusinessLogic.CrmIdValidationPolicy.Policy);
            // 
            // crmIdValidationLabel
            // 
            this.crmIdValidationLabel.AutoSize = true;
            this.crmIdValidationLabel.Location = new System.Drawing.Point(6, 163);
            this.crmIdValidationLabel.Name = "crmIdValidationLabel";
            this.crmIdValidationLabel.Size = new System.Drawing.Size(92, 13);
            this.crmIdValidationLabel.TabIndex = 37;
            this.crmIdValidationLabel.Text = "CRM Id Validation";
            // 
            // showErrorsSelector
            // 
            this.showErrorsSelector.DataSource = this.popupWhenBindingSource;
            this.showErrorsSelector.FormattingEnabled = true;
            this.showErrorsSelector.Location = new System.Drawing.Point(186, 136);
            this.showErrorsSelector.Name = "showErrorsSelector";
            this.showErrorsSelector.Size = new System.Drawing.Size(140, 21);
            this.showErrorsSelector.TabIndex = 36;
            // 
            // popupWhenBindingSource
            // 
            this.popupWhenBindingSource.DataSource = typeof(SuiteCRMAddIn.BusinessLogic.ErrorHandler.PopupWhen);
            // 
            // showErrorsLabel
            // 
            this.showErrorsLabel.AutoSize = true;
            this.showErrorsLabel.Location = new System.Drawing.Point(6, 139);
            this.showErrorsLabel.Name = "showErrorsLabel";
            this.showErrorsLabel.Size = new System.Drawing.Size(67, 13);
            this.showErrorsLabel.TabIndex = 35;
            this.showErrorsLabel.Text = "Show Errors:";
            // 
            // logLevelHelp
            // 
            this.logLevelHelp.AutoSize = true;
            this.logLevelHelp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.logLevelHelp.Location = new System.Drawing.Point(6, 109);
            this.logLevelHelp.Name = "logLevelHelp";
            this.logLevelHelp.Size = new System.Drawing.Size(307, 13);
            this.logLevelHelp.TabIndex = 34;
            this.logLevelHelp.Text = "Changes to log level do not take effect until you restart Outlook.";
            // 
            // logLevelLabel
            // 
            this.logLevelLabel.AutoSize = true;
            this.logLevelLabel.Location = new System.Drawing.Point(6, 84);
            this.logLevelLabel.Name = "logLevelLabel";
            this.logLevelLabel.Size = new System.Drawing.Size(53, 13);
            this.logLevelLabel.TabIndex = 33;
            this.logLevelLabel.Text = "Log level:";
            // 
            // logLevelSelector
            // 
            this.logLevelSelector.FormattingEnabled = true;
            this.logLevelSelector.Location = new System.Drawing.Point(186, 81);
            this.logLevelSelector.Name = "logLevelSelector";
            this.logLevelSelector.Size = new System.Drawing.Size(140, 21);
            this.logLevelSelector.TabIndex = 32;
            // 
            // LinkToLogFileDir
            // 
            this.LinkToLogFileDir.AutoEllipsis = true;
            this.LinkToLogFileDir.Location = new System.Drawing.Point(6, 54);
            this.LinkToLogFileDir.Name = "LinkToLogFileDir";
            this.LinkToLogFileDir.Size = new System.Drawing.Size(320, 14);
            this.LinkToLogFileDir.TabIndex = 3;
            this.LinkToLogFileDir.TabStop = true;
            this.LinkToLogFileDir.Text = "C:\\path\\to\\log\\files";
            this.LinkToLogFileDir.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkToLogFileDir_LinkClicked);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(6, 28);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(172, 13);
            this.label11.TabIndex = 4;
            this.label11.Text = "Log files are stored in this directory:";
            // 
            // AddInTitleLabel
            // 
            this.AddInTitleLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AddInTitleLabel.Location = new System.Drawing.Point(13, 26);
            this.AddInTitleLabel.Name = "AddInTitleLabel";
            this.AddInTitleLabel.Size = new System.Drawing.Size(332, 25);
            this.AddInTitleLabel.TabIndex = 1;
            this.AddInTitleLabel.Text = "SuiteCRM Outlook Add-In";
            this.AddInTitleLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(306, 390);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 42;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSave.Location = new System.Drawing.Point(225, 390);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 41;
            this.btnSave.Text = "&Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // startupDeferralGroup
            // 
            this.startupDeferralGroup.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.startupDeferralGroup.Controls.Add(this.startupDeferralHelp);
            this.startupDeferralGroup.Controls.Add(this.startupDeferralInput);
            this.startupDeferralGroup.Controls.Add(this.startupDeferralLabel);
            this.startupDeferralGroup.Location = new System.Drawing.Point(20, 182);
            this.startupDeferralGroup.Name = "startupDeferralGroup";
            this.startupDeferralGroup.Size = new System.Drawing.Size(334, 186);
            this.startupDeferralGroup.TabIndex = 9;
            this.startupDeferralGroup.TabStop = false;
            this.startupDeferralGroup.Text = "Startup Deferral";
            // 
            // startupDeferralLabel
            // 
            this.startupDeferralLabel.AutoSize = true;
            this.startupDeferralLabel.Location = new System.Drawing.Point(18, 22);
            this.startupDeferralLabel.Name = "startupDeferralLabel";
            this.startupDeferralLabel.Size = new System.Drawing.Size(118, 13);
            this.startupDeferralLabel.TabIndex = 11;
            this.startupDeferralLabel.Text = "Startup delay (seconds)";
            // 
            // startupDeferralInput
            // 
            this.startupDeferralInput.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.startupDeferralInput.Location = new System.Drawing.Point(142, 22);
            this.startupDeferralInput.Name = "startupDeferralInput";
            this.startupDeferralInput.Size = new System.Drawing.Size(186, 20);
            this.startupDeferralInput.TabIndex = 13;
            // 
            // startupDeferralHelp
            // 
            this.startupDeferralHelp.AutoSize = true;
            this.startupDeferralHelp.Location = new System.Drawing.Point(18, 63);
            this.startupDeferralHelp.Name = "startupDeferralHelp";
            this.startupDeferralHelp.Size = new System.Drawing.Size(0, 13);
            this.startupDeferralHelp.TabIndex = 14;
            // 
            // SettingsDialog
            // 
            this.AcceptButton = this.btnSave;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(393, 420);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(409, 458);
            this.Name = "SettingsDialog";
            this.Padding = new System.Windows.Forms.Padding(3);
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Settings";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmSettings_FormClosing);
            this.Load += new System.EventHandler(this.frmSettings_Load);
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.EmailArchiveAccountTabs.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            this.licenceGroup.ResumeLayout(false);
            this.licenceGroup.PerformLayout();
            this.gbLDAPAuthentication.ResumeLayout(false);
            this.gbLDAPAuthentication.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.synchronisationTab.ResumeLayout(false);
            this.synchronisationTab.PerformLayout();
            this.InformationTabPage.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.policyBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.popupWhenBindingSource)).EndInit();
            this.startupDeferralGroup.ResumeLayout(false);
            this.startupDeferralGroup.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.startupDeferralInput)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtAutoSync;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox checkBoxShowRightClick;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtSyncMaxRecords;
        private System.Windows.Forms.CheckBox cbShowCustomModules;
        private System.Windows.Forms.CheckBox checkBoxAutomaticSearch;
        private System.Windows.Forms.CheckBox cbEmailAttachments;
        private System.Windows.Forms.DateTimePicker dtpAutoArchiveFrom;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnTestLogin;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.TextBox txtUsername;
        private System.Windows.Forms.TextBox txtURL;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox gbLDAPAuthentication;
        private System.Windows.Forms.Label labelKey;
        private System.Windows.Forms.TextBox txtLDAPAuthenticationKey;
        private System.Windows.Forms.CheckBox chkEnableLDAPAuthentication;
        private System.Windows.Forms.Button btnSelect;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.CheckBox chkShowConfirmationMessageArchive;
        private System.Windows.Forms.TabPage InformationTabPage;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.LinkLabel LinkToLogFileDir;
        private System.Windows.Forms.Label AddInTitleLabel;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Label AddInVersionLabel;
        private System.Windows.Forms.TabControl EmailArchiveAccountTabs;
        private System.Windows.Forms.TabPage autoArchiveAccountsPage;
        private System.Windows.Forms.GroupBox licenceGroup;
        private System.Windows.Forms.Label licenceLabel;
        private System.Windows.Forms.TextBox licenceText;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label logLevelLabel;
        private System.Windows.Forms.ComboBox logLevelSelector;
        private System.Windows.Forms.Label logLevelHelp;
        private System.Windows.Forms.TabPage synchronisationTab;
        private System.Windows.Forms.ComboBox syncCallsMenu;
        private System.Windows.Forms.Label syncCallsLabel;
        private System.Windows.Forms.ComboBox syncContactsMenu;
        private System.Windows.Forms.Label syncContactsLabel;
        private System.Windows.Forms.Label syncLabel;
        private System.Windows.Forms.ComboBox syncMeetingsMenu;
        private System.Windows.Forms.Label syncMeetingsLabel;
        private System.Windows.Forms.ComboBox syncTasksMenu;
        private System.Windows.Forms.Label syncTasksLabel;
        private System.Windows.Forms.Button advancedArchiveSettingsButton;
        private System.Windows.Forms.ComboBox showErrorsSelector;
        private System.Windows.Forms.Label showErrorsLabel;
        private System.Windows.Forms.BindingSource popupWhenBindingSource;
        private System.Windows.Forms.ComboBox crmIdValidationSelector;
        private System.Windows.Forms.Label crmIdValidationLabel;
        private System.Windows.Forms.BindingSource policyBindingSource;
        private System.Windows.Forms.Label dotNetVersionLabel;
        private System.Windows.Forms.GroupBox startupDeferralGroup;
        private System.Windows.Forms.Label startupDeferralHelp;
        private System.Windows.Forms.NumericUpDown startupDeferralInput;
        private System.Windows.Forms.Label startupDeferralLabel;
    }
}
