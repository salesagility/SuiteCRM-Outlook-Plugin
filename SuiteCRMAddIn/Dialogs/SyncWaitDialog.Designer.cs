using System.Windows.Forms;

namespace SuiteCRMAddIn.Dialogs
{
    partial class SyncWaitDialog : Form
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SyncWaitDialog));
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.synchronisingMessageLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(15, 36);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(396, 23);
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar.TabIndex = 0;
            this.progressBar.UseWaitCursor = true;
            // 
            // synchronisingMessageLabel
            // 
            this.synchronisingMessageLabel.AutoEllipsis = true;
            this.synchronisingMessageLabel.AutoSize = true;
            this.synchronisingMessageLabel.Location = new System.Drawing.Point(12, 9);
            this.synchronisingMessageLabel.Name = "synchronisingMessageLabel";
            this.synchronisingMessageLabel.Size = new System.Drawing.Size(180, 13);
            this.synchronisingMessageLabel.TabIndex = 1;
            this.synchronisingMessageLabel.Text = "Please wait; synchronising with CRM";
            // 
            // SyncWaitDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(422, 76);
            this.Controls.Add(this.synchronisingMessageLabel);
            this.Controls.Add(this.progressBar);
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SyncWaitDialog";
            this.Text = "Please wait: synchronising with CRM";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label synchronisingMessageLabel;
    }
}