using SuiteCRMClient;
using SuiteCRMClient.Logging;
using SuiteCRMClient.RESTObjects;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SuiteCRMAddIn.Dialogs
{
    public partial class ArchiveModuleSelectFieldsDialog : Form
    {
        private readonly AvailableModules _availableModules;

        public ArchiveModuleSelectFieldsDialog()
        {
            InitializeComponent();

            _availableModules = RestAPIWrapper.GetModules();
            this.AddStandardModules();
        }

        /// <summary>
        /// Add the standard modules to the list view.
        /// </summary>
        private void AddStandardModules()
        {
            this.modulesSelector.Items.Clear();

            foreach (var module in ArchiveDialog.standardModules)
            {
                if (module != null)
                {
                    this.modulesSelector.Items.Add(module);
                }
            }
        }


        public ArchiveModuleSelectFieldsDialog(string text) : this()
        {
            var selectedIndex = modulesSelector.FindString(text);

            if (selectedIndex >= 0)
            {
                this.modulesSelector.SelectedItem = this.modulesSelector.Items[selectedIndex];
            }
        }

        private void moduleSelector_SelectionChanged(object sender, EventArgs e)
        {
            this.fieldsList.Items.Clear();

            foreach (var fieldName in new string[] {"Sample", "Field", "Names"})
            {
                fieldsList.Items.Add( new ListViewItem 
                {
                    Tag = fieldName,
                    Text = fieldName,
                    Checked = false
                });
            }
        }
    }
}
