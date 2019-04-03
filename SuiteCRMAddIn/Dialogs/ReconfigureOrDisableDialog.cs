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
    using System;
    using System.Windows.Forms;

    /// <summary>
    /// A simple dialog asking the user whether they wish to reconfigure or disable
    /// the add in following a licence check fail.
    /// </summary>
    public partial class ReconfigureOrDisableDialog : Form
    {
        /// <summary>
        /// Create the default 'Reconfigure or disable?' dialogue, with the heading
        /// 'Licence check failed'.
        /// </summary>
        public ReconfigureOrDisableDialog()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Create a 'Reconfigure or disable?' dialogue, with the heading and title
        /// taken from this summary.
        /// </summary>
        /// <param name="summary">Summary of the problem.</param>
        public ReconfigureOrDisableDialog(string summary, bool allowRetry) :this()
        {
            this.heading.Text = $"{summary} for the SuiteCRM add-in";
            this.Text = summary;
            this.RetryButton.Enabled = allowRetry;
        }

        private void DisableButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void ReconfigureButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
    }
}
