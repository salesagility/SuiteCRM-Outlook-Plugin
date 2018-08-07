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
    using SuiteCRMAddIn.BusinessLogic;
    partial class ArchiveDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ArchiveDialog));
            this.txtSearch = new System.Windows.Forms.TextBox();
            this.btnSearch = new System.Windows.Forms.Button();
            this.lstViewSearchModules = new System.Windows.Forms.ListView();
            this.colList = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.tsResults = new System.Windows.Forms.TreeView();
            this.btnArchive = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.categoryLabel = new System.Windows.Forms.Label();
            this.categoryInput = new System.Windows.Forms.ComboBox();
            this.legend = new System.Windows.Forms.TextBox();
            this.instructionLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtSearch
            // 
            this.txtSearch.Location = new System.Drawing.Point(9, 34);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(370, 20);
            this.txtSearch.TabIndex = 0;
            this.txtSearch.TextChanged += new System.EventHandler(this.txtSearch_TextChanged);
            this.txtSearch.Enter += new System.EventHandler(this.txtSearch_Enter);
            this.txtSearch.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSearch_KeyDown);
            // 
            // btnSearch
            // 
            this.btnSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSearch.Location = new System.Drawing.Point(385, 31);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(75, 23);
            this.btnSearch.TabIndex = 1;
            this.btnSearch.Text = "&Search";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // lstViewSearchModules
            // 
            this.lstViewSearchModules.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lstViewSearchModules.CheckBoxes = true;
            this.lstViewSearchModules.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colList});
            this.lstViewSearchModules.Location = new System.Drawing.Point(304, 60);
            this.lstViewSearchModules.Name = "lstViewSearchModules";
            this.lstViewSearchModules.Size = new System.Drawing.Size(156, 250);
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
            this.tsResults.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.tsResults.CheckBoxes = true;
            this.tsResults.Location = new System.Drawing.Point(9, 60);
            this.tsResults.Name = "tsResults";
            this.tsResults.Size = new System.Drawing.Size(289, 250);
            this.tsResults.TabIndex = 4;
            this.tsResults.Tag = "tree_search_results";
            this.tsResults.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.tsResults_AfterCheck);
            this.tsResults.AfterExpand += new System.Windows.Forms.TreeViewEventHandler(this.tsResults_AfterExpand);
            this.tsResults.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tsResults_NodeMouseClick);
            // 
            // btnArchive
            // 
            this.btnArchive.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnArchive.Enabled = false;
            this.btnArchive.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnArchive.Location = new System.Drawing.Point(304, 369);
            this.btnArchive.Name = "btnArchive";
            this.btnArchive.Size = new System.Drawing.Size(75, 23);
            this.btnArchive.TabIndex = 5;
            this.btnArchive.Text = "&Archive";
            this.btnArchive.UseVisualStyleBackColor = true;
            this.btnArchive.Click += new System.EventHandler(this.btnArchive_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(385, 369);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // categoryLabel
            // 
            this.categoryLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.categoryLabel.AutoSize = true;
            this.categoryLabel.Location = new System.Drawing.Point(9, 374);
            this.categoryLabel.Name = "categoryLabel";
            this.categoryLabel.Size = new System.Drawing.Size(49, 13);
            this.categoryLabel.TabIndex = 7;
            this.categoryLabel.Text = "Category";
            // 
            // categoryInput
            // 
            this.categoryInput.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.categoryInput.FormattingEnabled = true;
            this.categoryInput.Location = new System.Drawing.Point(68, 370);
            this.categoryInput.Name = "categoryInput";
            this.categoryInput.Size = new System.Drawing.Size(230, 21);
            this.categoryInput.TabIndex = 8;
            this.categoryInput.SelectedIndexChanged += new System.EventHandler(this.categoryInput_SelectedIndexChanged);
            // 
            // legend
            // 
            this.legend.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.legend.BackColor = System.Drawing.SystemColors.Control;
            this.legend.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.legend.ForeColor = System.Drawing.Color.DarkRed;
            this.legend.Location = new System.Drawing.Point(9, 316);
            this.legend.Multiline = true;
            this.legend.Name = "legend";
            this.legend.ReadOnly = true;
            this.legend.Size = new System.Drawing.Size(451, 47);
            this.legend.TabIndex = 9;
            this.legend.KeyUp += new System.Windows.Forms.KeyEventHandler(this.legend_KeyUp);
            this.legend.MouseUp += new System.Windows.Forms.MouseEventHandler(this.legend_MouseUp);
            // 
            // instructionLabel
            // 
            this.instructionLabel.AutoSize = true;
            this.instructionLabel.Location = new System.Drawing.Point(6, 9);
            this.instructionLabel.Name = "instructionLabel";
            this.instructionLabel.Size = new System.Drawing.Size(194, 13);
            this.instructionLabel.TabIndex = 10;
            this.instructionLabel.Text = "Use the form below find records in CRM";
            // 
            // ArchiveDialog
            // 
            this.AcceptButton = this.btnSearch;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(465, 399);
            this.Controls.Add(this.instructionLabel);
            this.Controls.Add(this.legend);
            this.Controls.Add(this.categoryInput);
            this.Controls.Add(this.categoryLabel);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnArchive);
            this.Controls.Add(this.tsResults);
            this.Controls.Add(this.lstViewSearchModules);
            this.Controls.Add(this.btnSearch);
            this.Controls.Add(this.txtSearch);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(481, 404);
            this.Name = "ArchiveDialog";
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
        private System.Windows.Forms.ListView lstViewSearchModules;
        private System.Windows.Forms.ColumnHeader colList;
        private System.Windows.Forms.TreeView tsResults;
        private System.Windows.Forms.Button btnArchive;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label categoryLabel;
        private System.Windows.Forms.ComboBox categoryInput;
        private System.Windows.Forms.TextBox legend;
        private System.Windows.Forms.Label instructionLabel;
    }
}
