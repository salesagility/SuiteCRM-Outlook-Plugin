namespace SuiteCRMAddIn.Dialogs
{
    partial class ConfirmRearchiveAlreadyArchivedEmails
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConfirmRearchiveAlreadyArchivedEmails));
            this.alreadyArchivedEmailsGrid = new System.Windows.Forms.DataGridView();
            this.From = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.To = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Subject = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.question = new System.Windows.Forms.Label();
            this.yesButton = new System.Windows.Forms.Button();
            this.noButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.alreadyArchivedEmailsGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // alreadyArchivedEmailsGrid
            // 
            this.alreadyArchivedEmailsGrid.AllowUserToAddRows = false;
            this.alreadyArchivedEmailsGrid.AllowUserToDeleteRows = false;
            this.alreadyArchivedEmailsGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.alreadyArchivedEmailsGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.alreadyArchivedEmailsGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.From,
            this.To,
            this.Date,
            this.Subject});
            this.alreadyArchivedEmailsGrid.Cursor = System.Windows.Forms.Cursors.Default;
            this.alreadyArchivedEmailsGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.alreadyArchivedEmailsGrid.Enabled = false;
            this.alreadyArchivedEmailsGrid.Location = new System.Drawing.Point(13, 25);
            this.alreadyArchivedEmailsGrid.Name = "alreadyArchivedEmailsGrid";
            this.alreadyArchivedEmailsGrid.ReadOnly = true;
            this.alreadyArchivedEmailsGrid.RowHeadersVisible = false;
            this.alreadyArchivedEmailsGrid.ShowEditingIcon = false;
            this.alreadyArchivedEmailsGrid.Size = new System.Drawing.Size(579, 146);
            this.alreadyArchivedEmailsGrid.TabIndex = 0;
            // 
            // From
            // 
            this.From.DataPropertyName = "From";
            this.From.HeaderText = "From";
            this.From.Name = "From";
            this.From.ReadOnly = true;
            // 
            // To
            // 
            this.To.DataPropertyName = "To";
            this.To.HeaderText = "To";
            this.To.Name = "To";
            this.To.ReadOnly = true;
            // 
            // Date
            // 
            this.Date.DataPropertyName = "Date";
            this.Date.HeaderText = "Date";
            this.Date.Name = "Date";
            this.Date.ReadOnly = true;
            // 
            // Subject
            // 
            this.Subject.DataPropertyName = "Subject";
            this.Subject.HeaderText = "Subject";
            this.Subject.Name = "Subject";
            this.Subject.ReadOnly = true;
            this.Subject.Width = 275;
            // 
            // question
            // 
            this.question.AutoSize = true;
            this.question.Location = new System.Drawing.Point(12, 9);
            this.question.Name = "question";
            this.question.Size = new System.Drawing.Size(442, 13);
            this.question.TabIndex = 1;
            this.question.Text = "The following emails have already been archived. Are you sure you want to re-arch" +
    "ive them?";
            // 
            // yesButton
            // 
            this.yesButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.yesButton.Cursor = System.Windows.Forms.Cursors.Default;
            this.yesButton.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.yesButton.Location = new System.Drawing.Point(377, 176);
            this.yesButton.Name = "yesButton";
            this.yesButton.Size = new System.Drawing.Size(64, 23);
            this.yesButton.TabIndex = 2;
            this.yesButton.Text = "Yes";
            this.yesButton.UseVisualStyleBackColor = true;
            // 
            // noButton
            // 
            this.noButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.noButton.DialogResult = System.Windows.Forms.DialogResult.No;
            this.noButton.Location = new System.Drawing.Point(447, 176);
            this.noButton.Name = "noButton";
            this.noButton.Size = new System.Drawing.Size(63, 23);
            this.noButton.TabIndex = 3;
            this.noButton.Text = "No";
            this.noButton.UseVisualStyleBackColor = true;
            // 
            // cancelButton
            // 
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(517, 176);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 4;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // ConfirmRearchiveAlreadyArchivedEmails
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.noButton;
            this.ClientSize = new System.Drawing.Size(604, 212);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.noButton);
            this.Controls.Add(this.yesButton);
            this.Controls.Add(this.question);
            this.Controls.Add(this.alreadyArchivedEmailsGrid);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(620, 250);
            this.Name = "ConfirmRearchiveAlreadyArchivedEmails";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Confirm Rearchive Already Archived Emails";
            ((System.ComponentModel.ISupportInitialize)(this.alreadyArchivedEmailsGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView alreadyArchivedEmailsGrid;
        private System.Windows.Forms.Label question;
        private System.Windows.Forms.Button yesButton;
        private System.Windows.Forms.Button noButton;
        private System.Windows.Forms.DataGridViewTextBoxColumn From;
        private System.Windows.Forms.DataGridViewTextBoxColumn To;
        private System.Windows.Forms.DataGridViewTextBoxColumn Date;
        private System.Windows.Forms.DataGridViewTextBoxColumn Subject;
        private System.Windows.Forms.Button cancelButton;
    }
}