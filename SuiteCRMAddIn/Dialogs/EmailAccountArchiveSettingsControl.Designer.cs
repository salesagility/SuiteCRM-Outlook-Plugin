using System.Windows.Forms;

namespace SuiteCRMAddIn.Dialogs
{
    partial class EmailAccountArchiveSettingsControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tsResults = new System.Windows.Forms.TreeView();
            this.LimitArchivingLabel = new System.Windows.Forms.Label();
            this.ArchiveOutboundCheckbox = new System.Windows.Forms.CheckBox();
            this.ArchiveInboundCheckbox = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // tsResults
            // 
            this.tsResults.AccessibleName = string.Empty;
            this.tsResults.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tsResults.CheckBoxes = true;
            this.tsResults.Location = new System.Drawing.Point(14, 78);
            this.tsResults.Name = "tsResults";
            this.tsResults.Size = new System.Drawing.Size(311, 139);
            this.tsResults.TabIndex = 15;
            this.tsResults.Tag = "tree_search_results";
            this.tsResults.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.tree_search_results_AfterCheck);
            this.tsResults.AfterExpand += new System.Windows.Forms.TreeViewEventHandler(this.tree_search_results_AfterExpand);
            this.tsResults.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tree_search_results_NodeMouseClick);
            // 
            // LimitArchivingLabel
            // 
            this.LimitArchivingLabel.AutoSize = true;
            this.LimitArchivingLabel.Location = new System.Drawing.Point(11, 62);
            this.LimitArchivingLabel.Name = "LimitArchivingLabel";
            this.LimitArchivingLabel.Size = new System.Drawing.Size(227, 13);
            this.LimitArchivingLabel.TabIndex = 14;
            this.LimitArchivingLabel.Text = "Archive All Messages In The Selected Folders:";
            // 
            // ArchiveOutboundCheckbox
            // 
            this.ArchiveOutboundCheckbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ArchiveOutboundCheckbox.AutoEllipsis = true;
            this.ArchiveOutboundCheckbox.Location = new System.Drawing.Point(11, 34);
            this.ArchiveOutboundCheckbox.Name = "ArchiveOutboundCheckbox";
            this.ArchiveOutboundCheckbox.Size = new System.Drawing.Size(314, 17);
            this.ArchiveOutboundCheckbox.TabIndex = 13;
            this.ArchiveOutboundCheckbox.Text = "Archive All Outbound (Sent) Messages";
            this.ArchiveOutboundCheckbox.UseVisualStyleBackColor = true;
            // 
            // ArchiveInboundCheckbox
            // 
            this.ArchiveInboundCheckbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ArchiveInboundCheckbox.AutoEllipsis = true;
            this.ArchiveInboundCheckbox.Location = new System.Drawing.Point(11, 11);
            this.ArchiveInboundCheckbox.Name = "ArchiveInboundCheckbox";
            this.ArchiveInboundCheckbox.Size = new System.Drawing.Size(314, 17);
            this.ArchiveInboundCheckbox.TabIndex = 16;
            this.ArchiveInboundCheckbox.Text = "Archive All Inbound (Received) Messages";
            this.ArchiveInboundCheckbox.UseVisualStyleBackColor = true;
            // 
            // EmailAccountArchiveSettingsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.Controls.Add(this.ArchiveInboundCheckbox);
            this.Controls.Add(this.tsResults);
            this.Controls.Add(this.LimitArchivingLabel);
            this.Controls.Add(this.ArchiveOutboundCheckbox);
            this.Name = "EmailAccountArchiveSettingsControl";
            this.Size = new System.Drawing.Size(332, 220);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TreeView tsResults;
        private System.Windows.Forms.Label LimitArchivingLabel;
        private System.Windows.Forms.CheckBox ArchiveOutboundCheckbox;
        private System.Windows.Forms.CheckBox ArchiveInboundCheckbox;
    }
}
