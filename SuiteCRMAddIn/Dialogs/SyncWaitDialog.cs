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
    using Outlook = Microsoft.Office.Interop.Outlook;
    using SuiteCRMAddIn.BusinessLogic;
    using SuiteCRMClient.Logging;
    using System.ComponentModel;
    using System.Threading;
    using System.Windows.Forms;

    /// <summary>
    /// A dialog to show while waiting to sync an item. At present this is necessarily a 
    /// meeting item, because other items are synced asynchronously through daemon actions.
    /// </summary>
    /// <remarks>
    /// Note that this shares a lot of commonality with 
    /// <see cref="Daemon.TransmitNewAction{Outlook.AppointmentItem}"/>, q.v.; some refactoring may
    /// be desirable.
    /// </remarks>
    public partial class SyncWaitDialog : Form 
    {
        /// <summary>
        /// the CRM type of the item I am sending.
        /// </summary>
        private readonly string crmType;

        /// <summary>
        ///  The fred in which things get done (don't know whether I actually need this)
        /// </summary>
        private BackgroundWorker fred = new BackgroundWorker();

        /// <summary>
        /// The logger through which I log.
        /// </summary>
        private readonly ILogger log;

        /// <summary>
        /// The synchroniser through which I synchronise.
        /// </summary>
        private Synchroniser<Microsoft.Office.Interop.Outlook.AppointmentItem> synchroniser;

        /// <summary>
        /// The state which wraps the item I am sending.
        /// </summary>
        private readonly SyncState<Outlook.AppointmentItem> syncState;

        public SyncWaitDialog(Synchroniser<Outlook.AppointmentItem> synchroniser, SyncState<Outlook.AppointmentItem> state, string crmType, ILogger log)
        {
            this.syncState = state;
            this.synchroniser = synchroniser;
            this.crmType = crmType;
            this.log = log;

            this.InitializeComponent();
            this.synchronisingMessageLabel.Text = $"Please wait: synchronising meeting '{state.OutlookItem.Subject}' with CRM";

            fred.RunWorkerCompleted += fred_Completed;
            fred.DoWork += fred_DoWork;

            fred.RunWorkerAsync();
        }

        private void fred_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            base.Close();
        }

        private void fred_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            if (Thread.CurrentThread.Name == null)
            {
                Thread.CurrentThread.Name = "Sync";
            }

            try
            {
                /* deal with any pending Windows messages, which we don't need to know about */
                Application.DoEvents();
                this.synchroniser.AddOrUpdateItemFromOutlookToCrm(syncState, this.crmType);
            }
            catch (System.Exception any)
            {
                log.Error($"Failure while syncing item", any);
            }
        }
    }
}
