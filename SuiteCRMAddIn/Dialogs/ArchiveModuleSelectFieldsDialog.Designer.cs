namespace SuiteCRMAddIn.Dialogs
{
    partial class ArchiveModuleSelectFieldsDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ArchiveModuleSelectFieldsDialog));
            this.fieldsToDisplayLabel = new System.Windows.Forms.Label();
            this.modulesSelector = new System.Windows.Forms.ComboBox();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.fieldsList = new System.Windows.Forms.ListView();
            this.fieldName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.valueType = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.SuspendLayout();
            // 
            // fieldsToDisplayLabel
            // 
            this.fieldsToDisplayLabel.AutoSize = true;
            this.fieldsToDisplayLabel.Location = new System.Drawing.Point(13, 13);
            this.fieldsToDisplayLabel.Name = "fieldsToDisplayLabel";
            this.fieldsToDisplayLabel.Size = new System.Drawing.Size(133, 13);
            this.fieldsToDisplayLabel.TabIndex = 0;
            this.fieldsToDisplayLabel.Text = "Fields to display for module";
            // 
            // modulesSelector
            // 
            this.modulesSelector.FormattingEnabled = true;
            this.modulesSelector.Location = new System.Drawing.Point(153, 13);
            this.modulesSelector.Name = "modulesSelector";
            this.modulesSelector.Size = new System.Drawing.Size(121, 21);
            this.modulesSelector.TabIndex = 1;
            this.modulesSelector.SelectedValueChanged += new System.EventHandler(this.moduleSelector_SelectionChanged);
            // 
            // okButton
            // 
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.okButton.Location = new System.Drawing.Point(118, 232);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 23);
            this.okButton.TabIndex = 3;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(199, 232);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 4;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // fieldsList
            // 
            this.fieldsList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.fieldsList.CheckBoxes = true;
            this.fieldsList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.fieldName,
            this.valueType});
            this.fieldsList.Location = new System.Drawing.Point(16, 40);
            this.fieldsList.Name = "fieldsList";
            this.fieldsList.Size = new System.Drawing.Size(256, 186);
            this.fieldsList.TabIndex = 5;
            this.fieldsList.UseCompatibleStateImageBehavior = false;
            this.fieldsList.View = System.Windows.Forms.View.Details;
            // 
            // fieldName
            // 
            this.fieldName.Text = "Field Name";
            this.fieldName.Width = 129;
            // 
            // valueType
            // 
            this.valueType.Text = "ValueType";
            this.valueType.Width = 122;
            // 
            // ArchiveModuleSelectFieldsDialog
            // 
            this.AcceptButton = this.okButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.fieldsList);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.modulesSelector);
            this.Controls.Add(this.fieldsToDisplayLabel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(300, 300);
            this.Name = "ArchiveModuleSelectFieldsDialog";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Select Fields to Display";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label fieldsToDisplayLabel;
        private System.Windows.Forms.ComboBox modulesSelector;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.ListView fieldsList;
        private System.Windows.Forms.ColumnHeader fieldName;
        private System.Windows.Forms.ColumnHeader valueType;
    }
}