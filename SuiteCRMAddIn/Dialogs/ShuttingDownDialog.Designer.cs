namespace SuiteCRMAddIn.Dialogs
{
    partial class ShuttingDownDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ShuttingDownDialog));
            this.infoLabel = new System.Windows.Forms.Label();
            this.progress = new System.Windows.Forms.ProgressBar();
            this.tasksRemainingLabel = new System.Windows.Forms.Label();
            this.icon = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.icon)).BeginInit();
            this.SuspendLayout();
            // 
            // infoLabel
            // 
            this.infoLabel.AutoSize = true;
            this.infoLabel.Location = new System.Drawing.Point(55, 9);
            this.infoLabel.Name = "infoLabel";
            this.infoLabel.Size = new System.Drawing.Size(227, 13);
            this.infoLabel.TabIndex = 0;
            this.infoLabel.Text = "Please wait: SuiteCRM Addin is shutting down.";
            // 
            // progress
            // 
            this.progress.Location = new System.Drawing.Point(58, 33);
            this.progress.Name = "progress";
            this.progress.Size = new System.Drawing.Size(456, 23);
            this.progress.TabIndex = 1;
            // 
            // tasksRemainingLabel
            // 
            this.tasksRemainingLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.tasksRemainingLabel.AutoSize = true;
            this.tasksRemainingLabel.Location = new System.Drawing.Point(385, 59);
            this.tasksRemainingLabel.Name = "tasksRemainingLabel";
            this.tasksRemainingLabel.Size = new System.Drawing.Size(124, 13);
            this.tasksRemainingLabel.TabIndex = 2;
            this.tasksRemainingLabel.Text = "100/100 tasks remaining";
            this.tasksRemainingLabel.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // icon
            // 
            this.icon.Image = global::SuiteCRMAddIn.Properties.Resources.SuiteCRMLogo;
            this.icon.Location = new System.Drawing.Point(12, 9);
            this.icon.Name = "icon";
            this.icon.Size = new System.Drawing.Size(37, 37);
            this.icon.TabIndex = 3;
            this.icon.TabStop = false;
            // 
            // ShuttingDownDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(527, 95);
            this.Controls.Add(this.icon);
            this.Controls.Add(this.tasksRemainingLabel);
            this.Controls.Add(this.progress);
            this.Controls.Add(this.infoLabel);
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ShuttingDownDialog";
            this.Text = "SuiteCRMAddin shutting down...";
            ((System.ComponentModel.ISupportInitialize)(this.icon)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label infoLabel;
        private System.Windows.Forms.ProgressBar progress;
        private System.Windows.Forms.Label tasksRemainingLabel;
        private System.Windows.Forms.PictureBox icon;
    }
}