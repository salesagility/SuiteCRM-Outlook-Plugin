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
    partial class frmAddressBook
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmAddressBook));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lstViewResults = new System.Windows.Forms.ListView();
            this.columnHeaderName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderEmail = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderModule = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnAddBCC = new System.Windows.Forms.Button();
            this.btnAddCC = new System.Windows.Forms.Button();
            this.btnRemoveBCC = new System.Windows.Forms.Button();
            this.btnRemoveCC = new System.Windows.Forms.Button();
            this.btnRemoveTo = new System.Windows.Forms.Button();
            this.btnAddTo = new System.Windows.Forms.Button();
            this.groupBoxBCC = new System.Windows.Forms.GroupBox();
            this.lstViewBCC = new System.Windows.Forms.ListView();
            this.columnHeader7 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader8 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader9 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.groupBoxCC = new System.Windows.Forms.GroupBox();
            this.lstViewCC = new System.Windows.Forms.ListView();
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader5 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader6 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.groupBoxTo = new System.Windows.Forms.GroupBox();
            this.lstViewTo = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.groupBoxSearch = new System.Windows.Forms.GroupBox();
            this.txtSearch = new System.Windows.Forms.TextBox();
            this.cbMyItems = new System.Windows.Forms.CheckBox();
            this.btnSearch = new System.Windows.Forms.Button();
            this.btnFinish = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBoxBCC.SuspendLayout();
            this.groupBoxCC.SuspendLayout();
            this.groupBoxTo.SuspendLayout();
            this.groupBoxSearch.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lstViewResults);
            this.groupBox1.Controls.Add(this.btnAddBCC);
            this.groupBox1.Controls.Add(this.btnAddCC);
            this.groupBox1.Controls.Add(this.btnRemoveBCC);
            this.groupBox1.Controls.Add(this.btnRemoveCC);
            this.groupBox1.Controls.Add(this.btnRemoveTo);
            this.groupBox1.Controls.Add(this.btnAddTo);
            this.groupBox1.Controls.Add(this.groupBoxBCC);
            this.groupBox1.Controls.Add(this.groupBoxCC);
            this.groupBox1.Controls.Add(this.groupBoxTo);
            this.groupBox1.Controls.Add(this.groupBoxSearch);
            this.groupBox1.Location = new System.Drawing.Point(6, 1);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(596, 357);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // lstViewResults
            // 
            this.lstViewResults.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderName,
            this.columnHeaderEmail,
            this.columnHeaderModule});
            this.lstViewResults.Location = new System.Drawing.Point(12, 93);
            this.lstViewResults.Name = "lstViewResults";
            this.lstViewResults.Size = new System.Drawing.Size(308, 234);
            this.lstViewResults.TabIndex = 1;
            this.lstViewResults.UseCompatibleStateImageBehavior = false;
            this.lstViewResults.View = System.Windows.Forms.View.Details;
            // 
            // columnHeaderName
            // 
            this.columnHeaderName.Text = "Name";
            this.columnHeaderName.Width = 130;
            // 
            // columnHeaderEmail
            // 
            this.columnHeaderEmail.Text = "Email";
            this.columnHeaderEmail.Width = 110;
            // 
            // columnHeaderModule
            // 
            this.columnHeaderModule.Text = "Type";
            this.columnHeaderModule.Width = 50;
            // 
            // btnAddBCC
            // 
            this.btnAddBCC.Font = new System.Drawing.Font("Tahoma", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAddBCC.ImageKey = "add.png";
            this.btnAddBCC.Location = new System.Drawing.Point(331, 275);
            this.btnAddBCC.Name = "btnAddBCC";
            this.btnAddBCC.Size = new System.Drawing.Size(50, 23);
            this.btnAddBCC.TabIndex = 28;
            this.btnAddBCC.Text = ">";
            this.btnAddBCC.UseVisualStyleBackColor = true;
            this.btnAddBCC.Click += new System.EventHandler(this.btnAddBCC_Click);
            // 
            // btnAddCC
            // 
            this.btnAddCC.Font = new System.Drawing.Font("Tahoma", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAddCC.ImageKey = "add.png";
            this.btnAddCC.Location = new System.Drawing.Point(331, 194);
            this.btnAddCC.Name = "btnAddCC";
            this.btnAddCC.Size = new System.Drawing.Size(50, 23);
            this.btnAddCC.TabIndex = 27;
            this.btnAddCC.Text = ">";
            this.btnAddCC.UseVisualStyleBackColor = true;
            this.btnAddCC.Click += new System.EventHandler(this.btnAddCC_Click);
            // 
            // btnRemoveBCC
            // 
            this.btnRemoveBCC.Font = new System.Drawing.Font("Tahoma", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRemoveBCC.ImageKey = "delete.png";
            this.btnRemoveBCC.Location = new System.Drawing.Point(331, 304);
            this.btnRemoveBCC.Name = "btnRemoveBCC";
            this.btnRemoveBCC.Size = new System.Drawing.Size(50, 23);
            this.btnRemoveBCC.TabIndex = 26;
            this.btnRemoveBCC.Text = "<";
            this.btnRemoveBCC.UseVisualStyleBackColor = true;
            this.btnRemoveBCC.Click += new System.EventHandler(this.btnRemoveBCC_Click);
            // 
            // btnRemoveCC
            // 
            this.btnRemoveCC.Font = new System.Drawing.Font("Tahoma", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRemoveCC.ImageKey = "delete.png";
            this.btnRemoveCC.Location = new System.Drawing.Point(331, 223);
            this.btnRemoveCC.Name = "btnRemoveCC";
            this.btnRemoveCC.Size = new System.Drawing.Size(50, 23);
            this.btnRemoveCC.TabIndex = 25;
            this.btnRemoveCC.Text = "<";
            this.btnRemoveCC.UseVisualStyleBackColor = true;
            this.btnRemoveCC.Click += new System.EventHandler(this.btnRemoveCC_Click);
            // 
            // btnRemoveTo
            // 
            this.btnRemoveTo.Font = new System.Drawing.Font("Tahoma", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRemoveTo.ImageKey = "delete.png";
            this.btnRemoveTo.Location = new System.Drawing.Point(331, 131);
            this.btnRemoveTo.Name = "btnRemoveTo";
            this.btnRemoveTo.Size = new System.Drawing.Size(50, 23);
            this.btnRemoveTo.TabIndex = 24;
            this.btnRemoveTo.Text = "<";
            this.btnRemoveTo.UseVisualStyleBackColor = true;
            this.btnRemoveTo.Click += new System.EventHandler(this.btnRemoveTo_Click);
            // 
            // btnAddTo
            // 
            this.btnAddTo.Font = new System.Drawing.Font("Tahoma", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAddTo.ImageKey = "add.png";
            this.btnAddTo.Location = new System.Drawing.Point(331, 102);
            this.btnAddTo.Name = "btnAddTo";
            this.btnAddTo.Size = new System.Drawing.Size(50, 23);
            this.btnAddTo.TabIndex = 23;
            this.btnAddTo.Text = ">";
            this.btnAddTo.UseVisualStyleBackColor = true;
            this.btnAddTo.Click += new System.EventHandler(this.btnAddTo_Click);
            // 
            // groupBoxBCC
            // 
            this.groupBoxBCC.Controls.Add(this.lstViewBCC);
            this.groupBoxBCC.Location = new System.Drawing.Point(387, 261);
            this.groupBoxBCC.Name = "groupBoxBCC";
            this.groupBoxBCC.Size = new System.Drawing.Size(200, 81);
            this.groupBoxBCC.TabIndex = 22;
            this.groupBoxBCC.TabStop = false;
            this.groupBoxBCC.Text = "BCC";
            // 
            // lstViewBCC
            // 
            this.lstViewBCC.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader7,
            this.columnHeader8,
            this.columnHeader9});
            this.lstViewBCC.Location = new System.Drawing.Point(6, 20);
            this.lstViewBCC.Name = "lstViewBCC";
            this.lstViewBCC.Size = new System.Drawing.Size(188, 52);
            this.lstViewBCC.TabIndex = 4;
            this.lstViewBCC.UseCompatibleStateImageBehavior = false;
            this.lstViewBCC.View = System.Windows.Forms.View.List;
            // 
            // columnHeader7
            // 
            this.columnHeader7.Text = "Name";
            this.columnHeader7.Width = 130;
            // 
            // columnHeader8
            // 
            this.columnHeader8.Text = "Email Address";
            this.columnHeader8.Width = 110;
            // 
            // columnHeader9
            // 
            this.columnHeader9.Text = "Module";
            this.columnHeader9.Width = 50;
            // 
            // groupBoxCC
            // 
            this.groupBoxCC.Controls.Add(this.lstViewCC);
            this.groupBoxCC.Location = new System.Drawing.Point(387, 172);
            this.groupBoxCC.Name = "groupBoxCC";
            this.groupBoxCC.Size = new System.Drawing.Size(200, 81);
            this.groupBoxCC.TabIndex = 21;
            this.groupBoxCC.TabStop = false;
            this.groupBoxCC.Text = "CC";
            // 
            // lstViewCC
            // 
            this.lstViewCC.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader4,
            this.columnHeader5,
            this.columnHeader6});
            this.lstViewCC.Location = new System.Drawing.Point(6, 22);
            this.lstViewCC.Name = "lstViewCC";
            this.lstViewCC.Size = new System.Drawing.Size(188, 52);
            this.lstViewCC.TabIndex = 3;
            this.lstViewCC.UseCompatibleStateImageBehavior = false;
            this.lstViewCC.View = System.Windows.Forms.View.List;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "Name";
            this.columnHeader4.Width = 130;
            // 
            // columnHeader5
            // 
            this.columnHeader5.Text = "Email Address";
            this.columnHeader5.Width = 110;
            // 
            // columnHeader6
            // 
            this.columnHeader6.Text = "Module";
            this.columnHeader6.Width = 50;
            // 
            // groupBoxTo
            // 
            this.groupBoxTo.Controls.Add(this.lstViewTo);
            this.groupBoxTo.Location = new System.Drawing.Point(387, 83);
            this.groupBoxTo.Name = "groupBoxTo";
            this.groupBoxTo.Size = new System.Drawing.Size(200, 81);
            this.groupBoxTo.TabIndex = 20;
            this.groupBoxTo.TabStop = false;
            this.groupBoxTo.Text = "To";
            // 
            // lstViewTo
            // 
            this.lstViewTo.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3});
            this.lstViewTo.Location = new System.Drawing.Point(6, 19);
            this.lstViewTo.Name = "lstViewTo";
            this.lstViewTo.Size = new System.Drawing.Size(188, 52);
            this.lstViewTo.TabIndex = 2;
            this.lstViewTo.UseCompatibleStateImageBehavior = false;
            this.lstViewTo.View = System.Windows.Forms.View.List;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Name";
            this.columnHeader1.Width = 130;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Email Address";
            this.columnHeader2.Width = 110;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Module";
            this.columnHeader3.Width = 50;
            // 
            // groupBoxSearch
            // 
            this.groupBoxSearch.Controls.Add(this.txtSearch);
            this.groupBoxSearch.Controls.Add(this.cbMyItems);
            this.groupBoxSearch.Controls.Add(this.btnSearch);
            this.groupBoxSearch.Location = new System.Drawing.Point(12, 22);
            this.groupBoxSearch.Name = "groupBoxSearch";
            this.groupBoxSearch.Size = new System.Drawing.Size(575, 54);
            this.groupBoxSearch.TabIndex = 17;
            this.groupBoxSearch.TabStop = false;
            this.groupBoxSearch.Text = "Search";
            // 
            // txtSearch
            // 
            this.txtSearch.Location = new System.Drawing.Point(15, 20);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(308, 20);
            this.txtSearch.TabIndex = 3;
            this.txtSearch.Enter += new System.EventHandler(this.txtSearch_Enter);
            this.txtSearch.Leave += new System.EventHandler(this.txtSearch_Leave);
            // 
            // cbMyItems
            // 
            this.cbMyItems.AutoSize = true;
            this.cbMyItems.Location = new System.Drawing.Point(343, 23);
            this.cbMyItems.Name = "cbMyItems";
            this.cbMyItems.Size = new System.Drawing.Size(90, 17);
            this.cbMyItems.TabIndex = 2;
            this.cbMyItems.Text = "Only my items";
            this.cbMyItems.UseVisualStyleBackColor = true;
            // 
            // btnSearch
            // 
            this.btnSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSearch.Location = new System.Drawing.Point(450, 18);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(75, 23);
            this.btnSearch.TabIndex = 1;
            this.btnSearch.Text = "&Search";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // btnFinish
            // 
            this.btnFinish.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFinish.Location = new System.Drawing.Point(446, 364);
            this.btnFinish.Name = "btnFinish";
            this.btnFinish.Size = new System.Drawing.Size(75, 23);
            this.btnFinish.TabIndex = 1;
            this.btnFinish.Text = "&Finish";
            this.btnFinish.UseVisualStyleBackColor = true;
            this.btnFinish.Click += new System.EventHandler(this.btnFinish_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(527, 364);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // frmAddressBook
            // 
            this.AcceptButton = this.btnFinish;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(608, 393);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnFinish);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(624, 431);
            this.Name = "frmAddressBook";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SuiteCRM Address Book";
            this.Load += new System.EventHandler(this.frmAddressBook_Load);
            this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.frmAddressBook_KeyPress);
            this.groupBox1.ResumeLayout(false);
            this.groupBoxBCC.ResumeLayout(false);
            this.groupBoxCC.ResumeLayout(false);
            this.groupBoxTo.ResumeLayout(false);
            this.groupBoxSearch.ResumeLayout(false);
            this.groupBoxSearch.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnAddBCC;
        private System.Windows.Forms.Button btnAddCC;
        private System.Windows.Forms.Button btnRemoveBCC;
        private System.Windows.Forms.Button btnRemoveCC;
        private System.Windows.Forms.Button btnRemoveTo;
        private System.Windows.Forms.Button btnAddTo;
        private System.Windows.Forms.GroupBox groupBoxBCC;
        private System.Windows.Forms.ListView lstViewBCC;
        private System.Windows.Forms.ColumnHeader columnHeader7;
        private System.Windows.Forms.ColumnHeader columnHeader8;
        private System.Windows.Forms.ColumnHeader columnHeader9;
        private System.Windows.Forms.GroupBox groupBoxCC;
        private System.Windows.Forms.ListView lstViewCC;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.ColumnHeader columnHeader5;
        private System.Windows.Forms.ColumnHeader columnHeader6;
        private System.Windows.Forms.GroupBox groupBoxTo;
        private System.Windows.Forms.ListView lstViewTo;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ListView lstViewResults;
        private System.Windows.Forms.ColumnHeader columnHeaderName;
        private System.Windows.Forms.ColumnHeader columnHeaderEmail;
        private System.Windows.Forms.ColumnHeader columnHeaderModule;
        private System.Windows.Forms.GroupBox groupBoxSearch;
        private System.Windows.Forms.CheckBox cbMyItems;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.Button btnFinish;
        private System.Windows.Forms.Button btnCancel;
    }
}