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
    using Newtonsoft.Json.Linq;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Linq;
    using System.Windows.Forms;

    public partial class AdvancedArchiveSettingsDialog : Form
    {
        private string backup;

        public AdvancedArchiveSettingsDialog()
        {
            InitializeComponent();
            this.backup = Properties.Settings.Default.ArchivingSearchChains;
            this.archiveSearchChainsText.Text = Properties.Settings.Default.ArchivingSearchChains;
        }

        internal static IDictionary<string, ICollection<ArchiveDialog.LinkSpec>> SetupSearchChains()
        {
            Dictionary<string, ICollection<ArchiveDialog.LinkSpec>> moduleChains = new Dictionary<string, ICollection<ArchiveDialog.LinkSpec>>();

            if (Properties.Settings.Default.ArchivingSearchChains != null)
            {
                JObject parsed = JObject.Parse(Properties.Settings.Default.ArchivingSearchChains);

                foreach (string key in parsed.Properties().Select(p => p.Name).ToList())
                {
                    List<ArchiveDialog.LinkSpec> links = new List<ArchiveDialog.LinkSpec>();

                    foreach (JObject parsedLink in parsed[key])
                    {
                        ICollection<string> fields = parsedLink["fields"].Select(f => f.ToString()).ToList();

                        links.Add(new ArchiveDialog.LinkSpec(parsedLink["linkName"].ToString(), parsedLink["targetName"].ToString(), fields));
                    }

                    moduleChains[key] = links;
                }
            }

            return moduleChains;
        }

        private void cancelButton_click(object sender, EventArgs e)
        {
            Properties.Settings.Default.ArchivingSearchChains = this.backup;
            this.Close();
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.UseWaitCursor = true;
                Properties.Settings.Default.ArchivingSearchChains = this.archiveSearchChainsText.Text;

                SetupSearchChains();
                this.Close();
            }
            catch (Exception any)
            {
                MessageBox.Show(any.Message, "JSON Parse Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.UseWaitCursor = false;
            }
        }
    }
}
