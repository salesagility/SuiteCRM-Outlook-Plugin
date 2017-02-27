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
namespace SuiteCRMAddIn.BusinessLogic
{
    using SuiteCRMClient.Logging;
    using System;
    using System.Threading;
    using System.Threading.Tasks;

    /// <summary>
    /// Do something repeatedly every five minutes.
    /// </summary>
    public abstract class RepeatingProcess
    {
        /// <summary>
        /// The polling interval; currently hard wired.
        /// </summary>
        private readonly TimeSpan SyncPeriod = TimeSpan.FromMinutes(5);

        /// <summary>
        /// The thread in which syncing is run.
        /// </summary>
        private Thread process;

        /// <summary>
        /// The logger to which I log.
        /// </summary>
        protected readonly ILogger Log;

        /// <summary>
        /// The run state I am currently in.
        /// </summary>
        private RunState state = RunState.Stopped;

        /// <summary>
        /// The name by which I am known.
        /// </summary>
        private readonly string Name;

        public RepeatingProcess(string name, ILogger log)
        {
            this.Log = log;
            this.Name = name;
            this.process = new Thread(() => this.PerformRepeatedly());
            this.process.Name = $"{this.Name}";
        }

        /// <summary>
        /// True if I should be running, else false.
        /// </summary>
        private Boolean Running
        {
            get { return this.state == RunState.Running; }
        }

        /// <summary>
        /// Do whatever it is I do repeatedly.
        /// </summary>
        private async void PerformRepeatedly()
        {
            while (this.Running)
            {
                Robustness.DoOrLogError(
                    this.Log, 
                    () => this.PerformIteration(), 
                    $"{this.GetType().Name} PerformIteration");

                await Task.Delay(this.SyncPeriod);
            }

            this.state = RunState.Stopped;
        }

        /// <summary>
        /// Do whatever it is I do, once.
        /// </summary>
        internal abstract void PerformIteration();

        /// <summary>
        /// Stop me at the end of my current iteration; does not force an immediate stop.
        /// </summary>
        public void Stop()
        {
            Log.Debug($"Stopping thread Sync{this.GetType().Name} at end of current iteration");
            this.state = RunState.Stopping;
        }

        /// <summary>
        /// If I am not currently running, set me running.
        /// </summary>
        public void Start()
        {
            if (!this.Running)
            {
                Log.Debug($"Starting thread Sync{this.GetType().Name}");
                this.state = RunState.Running;
                this.process.Start();
            }
            else
            {
                Log.Warn($"Did not start thread Sync{this.GetType().Name} as it appears to be running");
            }
        }
    }
}