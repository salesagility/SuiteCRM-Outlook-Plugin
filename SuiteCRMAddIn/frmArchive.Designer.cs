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
    partial class frmArchive
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
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem("All");
            System.Windows.Forms.ListViewItem listViewItem2 = new System.Windows.Forms.ListViewItem("Accounts");
            System.Windows.Forms.ListViewItem listViewItem3 = new System.Windows.Forms.ListViewItem("Contacts");
            System.Windows.Forms.ListViewItem listViewItem4 = new System.Windows.Forms.ListViewItem("Leads");
            System.Windows.Forms.ListViewItem listViewItem5 = new System.Windows.Forms.ListViewItem("Opportunities");
            System.Windows.Forms.ListViewItem listViewItem6 = new System.Windows.Forms.ListViewItem("Cases");
            System.Windows.Forms.ListViewItem listViewItem7 = new System.Windows.Forms.ListViewItem("Project");
            System.Windows.Forms.ListViewItem listViewItem8 = new System.Windows.Forms.ListViewItem("Bugs");
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmArchive));
            this.txtSearch = new System.Windows.Forms.TextBox();
            this.btnSearch = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lstViewSearchModules = new System.Windows.Forms.ListView();
            this.colList = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.tsResults = new System.Windows.Forms.TreeView();
            this.btnArchive = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtSearch
            // 
            this.txtSearch.Location = new System.Drawing.Point(9, 30);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(370, 20);
            this.txtSearch.TabIndex = 0;
            this.txtSearch.TextChanged += new System.EventHandler(this.txtSearch_TextChanged);
            this.txtSearch.Enter += new System.EventHandler(this.txtSearch_Enter);
            this.txtSearch.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSearch_KeyDown);
            this.txtSearch.Leave += new System.EventHandler(this.txtSearch_Leave);
            // 
            // btnSearch
            // 
            this.btnSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSearch.Location = new System.Drawing.Point(385, 28);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(75, 23);
            this.btnSearch.TabIndex = 1;
            this.btnSearch.Text = "&Search";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(230, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Use the form below to find records in SuiteCRM";
            // 
            // lstViewSearchModules
            // 
            this.lstViewSearchModules.CheckBoxes = true;
            this.lstViewSearchModules.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colList});
            listViewItem1.StateImageIndex = 0;
            listViewItem1.Tag = "All";
            listViewItem2.StateImageIndex = 0;
            listViewItem2.Tag = "Accounts";
            listViewItem3.StateImageIndex = 0;
            listViewItem3.Tag = "Contacts";
            listViewItem4.StateImageIndex = 0;
            listViewItem4.Tag = "Leads";
            listViewItem5.StateImageIndex = 0;
            listViewItem5.Tag = "Opportunities";
            listViewItem6.StateImageIndex = 0;
            listViewItem6.Tag = "Cases";
            listViewItem7.StateImageIndex = 0;
            listViewItem7.Tag = "Project";
            listViewItem8.StateImageIndex = 0;
            listViewItem8.Tag = "Bugs";
            this.lstViewSearchModules.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1,
            listViewItem2,
            listViewItem3,
            listViewItem4,
            listViewItem5,
            listViewItem6,
            listViewItem7,
            listViewItem8});
            this.lstViewSearchModules.Location = new System.Drawing.Point(326, 68);
            this.lstViewSearchModules.Name = "lstViewSearchModules";
            this.lstViewSearchModules.Size = new System.Drawing.Size(134, 262);
            this.lstViewSearchModules.TabIndex = 3;
            this.lstViewSearchModules.UseCompatibleStateImageBehavior = false;
            this.lstViewSearchModules.View = System.Windows.Forms.View.Details;
            this.lstViewSearchModules.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.lstViewSearchModules_ItemChecked);
            // 
            // colList
            // 
            this.colList.Text = "Modules";
            this.colList.Width = 129;
            // 
            // tsResults
            // 
            this.tsResults.AccessibleName = "";
            this.tsResults.CheckBoxes = true;
            this.tsResults.Location = new System.Drawing.Point(9, 68);
            this.tsResults.Name = "tsResults";
            this.tsResults.Size = new System.Drawing.Size(311, 262);
            this.tsResults.TabIndex = 4;
            this.tsResults.Tag = "tree_search_results";
            this.tsResults.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.tsResults_AfterCheck);
            this.tsResults.AfterExpand += new System.Windows.Forms.TreeViewEventHandler(this.tsResults_AfterExpand);
            this.tsResults.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tsResults_NodeMouseClick);
            // 
            // btnArchive
            // 
            this.btnArchive.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnArchive.Location = new System.Drawing.Point(304, 336);
            this.btnArchive.Name = "btnArchive";
            this.btnArchive.Size = new System.Drawing.Size(75, 23);
            this.btnArchive.TabIndex = 5;
            this.btnArchive.Text = "&Archive";
            this.btnArchive.UseVisualStyleBackColor = true;
            this.btnArchive.Click += new System.EventHandler(this.btnArchive_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(385, 336);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // frmArchive
            // 
            this.AcceptButton = this.btnArchive;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(465, 366);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnArchive);
            this.Controls.Add(this.tsResults);
            this.Controls.Add(this.lstViewSearchModules);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnSearch);
            this.Controls.Add(this.txtSearch);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(481, 404);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(481, 404);
            this.Name = "frmArchive";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Archive to SuiteCRM";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmArchive_FormClosed);
            this.Load += new System.EventHandler(this.frmArchive_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public  System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListView lstViewSearchModules;
        private System.Windows.Forms.ColumnHeader colList;
        private System.Windows.Forms.TreeView tsResults;
        private System.Windows.Forms.Button btnArchive;
        private System.Windows.Forms.Button btnCancel;      
    }
}