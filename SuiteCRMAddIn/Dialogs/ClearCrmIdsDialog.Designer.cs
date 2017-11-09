namespace SuiteCRMAddIn.Dialogs
{
    partial class ClearCrmIdsDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ClearCrmIdsDialog));
            this.question = new System.Windows.Forms.Label();
            this.icon = new System.Windows.Forms.PictureBox();
            this.noButton = new System.Windows.Forms.Button();
            this.yesButton = new System.Windows.Forms.Button();
            this.progress = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.icon)).BeginInit();
            this.SuspendLayout();
            // 
            // question
            // 
            this.question.AutoSize = true;
            this.question.Location = new System.Drawing.Point(55, 12);
            this.question.Name = "question";
            this.question.Size = new System.Drawing.Size(372, 13);
            this.question.TabIndex = 0;
            this.question.Text = "You have changed the CRM URL. Do you want to clear all existing CRM ids?";
            // 
            // icon
            // 
            this.icon.Image = global::SuiteCRMAddIn.Properties.Resources.SuiteCRMLogo;
            this.icon.Location = new System.Drawing.Point(12, 12);
            this.icon.Name = "icon";
            this.icon.Size = new System.Drawing.Size(37, 37);
            this.icon.TabIndex = 4;
            this.icon.TabStop = false;
            // 
            // noButton
            // 
            this.noButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.noButton.DialogResult = System.Windows.Forms.DialogResult.No;
            this.noButton.Location = new System.Drawing.Point(552, 37);
            this.noButton.Name = "noButton";
            this.noButton.Size = new System.Drawing.Size(75, 23);
            this.noButton.TabIndex = 2;
            this.noButton.Text = "No";
            this.noButton.UseVisualStyleBackColor = true;
            // 
            // yesButton
            // 
            this.yesButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.yesButton.Location = new System.Drawing.Point(471, 37);
            this.yesButton.Name = "yesButton";
            this.yesButton.Size = new System.Drawing.Size(75, 23);
            this.yesButton.TabIndex = 1;
            this.yesButton.Text = "Yes";
            this.yesButton.UseVisualStyleBackColor = true;
            this.yesButton.Click += new System.EventHandler(this.yesButton_Click);
            // 
            // progress
            // 
            this.progress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progress.Location = new System.Drawing.Point(58, 37);
            this.progress.Name = "progress";
            this.progress.Size = new System.Drawing.Size(407, 23);
            this.progress.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progress.TabIndex = 3;
            this.progress.Visible = false;
            // 
            // ClearCrmIdsDialog
            // 
            this.AcceptButton = this.yesButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(639, 72);
            this.Controls.Add(this.progress);
            this.Controls.Add(this.yesButton);
            this.Controls.Add(this.noButton);
            this.Controls.Add(this.icon);
            this.Controls.Add(this.question);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ClearCrmIdsDialog";
            this.Text = "ClearCrmIds";
            ((System.ComponentModel.ISupportInitialize)(this.icon)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label question;
        private System.Windows.Forms.PictureBox icon;
        private System.Windows.Forms.Button noButton;
        private System.Windows.Forms.Button yesButton;
        private System.Windows.Forms.ProgressBar progress;
    }
}