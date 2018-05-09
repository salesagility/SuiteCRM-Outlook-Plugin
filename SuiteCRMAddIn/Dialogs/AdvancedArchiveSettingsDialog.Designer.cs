namespace SuiteCRMAddIn.Dialogs
{
    partial class AdvancedArchiveSettingsDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AdvancedArchiveSettingsDialog));
            this.archivingSearchChainsButton = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.archiveSearchChainsText = new System.Windows.Forms.TextBox();
            this.parseFeedback = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // archivingSearchChainsButton
            // 
            this.archivingSearchChainsButton.AutoSize = true;
            this.archivingSearchChainsButton.Location = new System.Drawing.Point(12, 9);
            this.archivingSearchChainsButton.Name = "archivingSearchChainsButton";
            this.archivingSearchChainsButton.Size = new System.Drawing.Size(160, 13);
            this.archivingSearchChainsButton.TabIndex = 0;
            this.archivingSearchChainsButton.Text = "Archiving Search Chains (JSON)";
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(497, 127);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 44;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.cancelButton_click);
            // 
            // btnSave
            // 
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSave.Location = new System.Drawing.Point(416, 127);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 43;
            this.btnSave.Text = "&Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.saveButton_Click);
            // 
            // archiveSearchChainsText
            // 
            this.archiveSearchChainsText.AcceptsReturn = true;
            this.archiveSearchChainsText.AcceptsTab = true;
            this.archiveSearchChainsText.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.archiveSearchChainsText.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.archiveSearchChainsText.Location = new System.Drawing.Point(15, 26);
            this.archiveSearchChainsText.Multiline = true;
            this.archiveSearchChainsText.Name = "archiveSearchChainsText";
            this.archiveSearchChainsText.Size = new System.Drawing.Size(557, 95);
            this.archiveSearchChainsText.TabIndex = 45;
            // 
            // parseFeedback
            // 
            this.parseFeedback.AutoSize = true;
            this.parseFeedback.Location = new System.Drawing.Point(12, 133);
            this.parseFeedback.Name = "parseFeedback";
            this.parseFeedback.Size = new System.Drawing.Size(0, 13);
            this.parseFeedback.TabIndex = 46;
            // 
            // AdvancedArchiveSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(584, 162);
            this.Controls.Add(this.parseFeedback);
            this.Controls.Add(this.archiveSearchChainsText);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.archivingSearchChainsButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(600, 200);
            this.MinimizeBox = false;
            this.Name = "AdvancedArchiveSettings";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Advanced Archive Settings";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label archivingSearchChainsButton;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox archiveSearchChainsText;
        private System.Windows.Forms.Label parseFeedback;
    }
}