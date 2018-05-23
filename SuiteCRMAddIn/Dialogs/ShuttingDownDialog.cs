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
    using SuiteCRMAddIn.BusinessLogic;
    using SuiteCRMClient.Logging;
    using System.ComponentModel;
    using System.Threading;
    using System.Windows.Forms;

    public partial class ShuttingDownDialog : Form
    {
        /// <summary>
        ///  The fred in which things get done (don't know whether I actually need this)
        /// </summary>
        static BackgroundWorker fred = new BackgroundWorker();

        /// <summary>
        /// The total number of tasks that remained to do when I started.
        /// </summary>
        private int tasks;

        /// <summary>
        /// The number of tasks remaining.
        /// </summary>
        private int remaining;

        /// <summary>
        /// The logger to which I shall log.
        /// </summary>
        private ILogger log;

        /// <summary>
        /// Create a new instance of a ShuttingDownDialog with this number of tasks to complete.
        /// </summary>
        /// <param name="tasks">the number of tasks to complete.</param>
        public ShuttingDownDialog(int tasks, ILogger log)
        {
            this.tasks = tasks;
            this.remaining = tasks;
            this.log = log;

            InitializeComponent();
            this.showProgressOrClose();

            fred.WorkerReportsProgress = true;
            
            fred.ProgressChanged += fred_ProgressChanged;
            fred.DoWork += fred_DoWork;

            fred.RunWorkerAsync();
        }

        private void fred_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            if (Thread.CurrentThread.Name == null)
            {
                Thread.CurrentThread.Name = "Shutdown";
            }

            if (worker != null)
            {
                while (remaining > 0)
                {
                    this.remaining = RepeatingProcess.PrepareShutdownAll(this.log);
                    double percentageRemaining = (100.0 * this.remaining) / this.tasks;
                    worker.ReportProgress((int)(100.0 - percentageRemaining));

                    /* deal with any pending Windows messages, which we don't need to know about */
                    Application.DoEvents();

                    Thread.Sleep(1000);
                }
            }
        }

        /// <summary>
        /// Receive the progress changed event in the user interface thread.
        /// </summary>
        /// <param name="sender">Fred</param>
        /// <param name="e">The event (which contains the percentage but in fact we don't use it).</param>
        private void fred_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.showProgressOrClose();
        }

        /// <summary>
        /// Update my progress indicators; if I am finished, close.
        /// </summary>
        private void showProgressOrClose()
        {
            if (this.remaining == 0)
            {
                this.log.Debug("ShuttingDownDialog: completed, closing.");
                base.Close();
            }
            else
            {
                double percentageRemaining = (100.0 * this.remaining) / this.tasks;

                this.progress.Value = 100 - (int)percentageRemaining;
                this.tasksRemainingLabel.Text = $"{this.remaining}/{this.tasks} tasks remaining";
                this.log.Debug($"ShuttingDownDialog: progress {percentageRemaining}%; {this.remaining}/{this.tasks}");
            }
        }
    }
}
