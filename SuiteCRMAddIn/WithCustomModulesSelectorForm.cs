using SuiteCRMClient;
using SuiteCRMClient.RESTObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SuiteCRMAddIn
{
    /// <summary>
    /// Both frmCustomModules and frmArchive user a custom modules selector, which each was building
    /// in a different way and each was building badly. This is a common superclass for both forms,
    /// adding a single simple populator for this control.
    /// </summary>
    public abstract class WithCustomModulesSelectorForm : Form
    {
        protected void PopulateCustomModulesListView(ListView view, List<string> ignoreModules)
        {
            foreach (module_data module in clsSuiteCRMHelper.GetModules().items.OrderBy(i => i.module_key))
            {
                string moduleKey = module.module_key;
                if (!ignoreModules.Contains(moduleKey))
                {
                    ListViewItem item = new ListViewItem
                    {
                        Checked = Globals.ThisAddIn.Settings.CustomModules.Select(i => i == moduleKey).Count() > 0,
                        Text = moduleKey,
                        Tag = moduleKey,
                        SubItems = { module.module_label }
                    };

                    if (view.Items.Cast<ListViewItem>().Select(i => i.Text == moduleKey).Count() > 0)
                    {
                        view.Items.Add(item);
                    }
                }
            }
        }
    }
}
