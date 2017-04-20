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
namespace SuiteCRMAddIn
{
    using SuiteCRMClient;
    using SuiteCRMClient.RESTObjects;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows.Forms;

    /// <summary>
    /// Both frmCustomModules and frmArchive user a custom modules selector, which each was building
    /// in a different way and each was building badly. This is a common superclass for both forms,
    /// providing a single simple populator for this control.
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
