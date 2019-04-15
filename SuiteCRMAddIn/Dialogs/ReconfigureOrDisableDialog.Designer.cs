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
    partial class ReconfigureOrDisableDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReconfigureOrDisableDialog));
            this.icon = new System.Windows.Forms.PictureBox();
            this.heading = new System.Windows.Forms.Label();
            this.question = new System.Windows.Forms.TextBox();
            this.DisableButton = new System.Windows.Forms.Button();
            this.ReconfigureButton = new System.Windows.Forms.Button();
            this.RetryButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.icon)).BeginInit();
            this.SuspendLayout();
            // 
            // icon
            // 
            this.icon.Image = global::SuiteCRMAddIn.Properties.Resources.SuiteCRMLogo;
            this.icon.Location = new System.Drawing.Point(12, 12);
            this.icon.Name = "icon";
            this.icon.Size = new System.Drawing.Size(37, 37);
            this.icon.TabIndex = 0;
            this.icon.TabStop = false;
            // 
            // heading
            // 
            this.heading.AutoSize = true;
            this.heading.Location = new System.Drawing.Point(77, 12);
            this.heading.Name = "heading";
            this.heading.Size = new System.Drawing.Size(243, 13);
            this.heading.TabIndex = 0;
            this.heading.Text = "The licence check for the SuiteCRM add-in failed.";
            // 
            // question
            // 
            this.question.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.question.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.question.Enabled = false;
            this.question.Location = new System.Drawing.Point(80, 36);
            this.question.Multiline = true;
            this.question.Name = "question";
            this.question.Size = new System.Drawing.Size(450, 42);
            this.question.TabIndex = 0;
            this.question.TabStop = false;
            this.question.Text = "Would you like to reconfigure the add-in and try again, or disable the add-in and" +
    " use Outlook without it?";
            // 
            // DisableButton
            // 
            this.DisableButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.DisableButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.DisableButton.Location = new System.Drawing.Point(455, 85);
            this.DisableButton.Name = "DisableButton";
            this.DisableButton.Size = new System.Drawing.Size(75, 23);
            this.DisableButton.TabIndex = 2;
            this.DisableButton.Text = "Disable";
            this.DisableButton.UseVisualStyleBackColor = true;
            this.DisableButton.Click += new System.EventHandler(this.DisableButton_Click);
            // 
            // ReconfigureButton
            // 
            this.ReconfigureButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.ReconfigureButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.ReconfigureButton.Location = new System.Drawing.Point(374, 84);
            this.ReconfigureButton.Name = "ReconfigureButton";
            this.ReconfigureButton.Size = new System.Drawing.Size(75, 23);
            this.ReconfigureButton.TabIndex = 1;
            this.ReconfigureButton.Text = "Reconfigure";
            this.ReconfigureButton.UseVisualStyleBackColor = true;
            this.ReconfigureButton.Click += new System.EventHandler(this.ReconfigureButton_Click);
            // 
            // RetryButton
            // 
            this.RetryButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.RetryButton.DialogResult = System.Windows.Forms.DialogResult.Retry;
            this.RetryButton.Location = new System.Drawing.Point(293, 85);
            this.RetryButton.Name = "RetryButton";
            this.RetryButton.Size = new System.Drawing.Size(75, 23);
            this.RetryButton.TabIndex = 3;
            this.RetryButton.Text = "Retry";
            this.RetryButton.UseVisualStyleBackColor = true;
            // 
            // ReconfigureOrDisableDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(542, 116);
            this.Controls.Add(this.RetryButton);
            this.Controls.Add(this.ReconfigureButton);
            this.Controls.Add(this.DisableButton);
            this.Controls.Add(this.question);
            this.Controls.Add(this.heading);
            this.Controls.Add(this.icon);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ReconfigureOrDisableDialog";
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Licence check failed";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.icon)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox icon;
        private System.Windows.Forms.Label heading;
        private System.Windows.Forms.TextBox question;
        private System.Windows.Forms.Button DisableButton;
        private System.Windows.Forms.Button ReconfigureButton;
        private System.Windows.Forms.Button RetryButton;
    }
}