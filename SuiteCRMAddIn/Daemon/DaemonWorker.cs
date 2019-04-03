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
namespace SuiteCRMAddIn.Daemon
{
    using BusinessLogic;
    using Exceptions;
    using SuiteCRMClient.Logging;
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// A thing which maintains a queue of tasks (instances of DaemonAction) and executes them in turn. Explicitly
    /// a singleton.
    /// </summary>
    public class DaemonWorker : RepeatingProcess
    {
        /// <summary>
        /// My underlying instance.
        /// </summary>
        private static readonly Lazy<DaemonWorker> lazy =
            new Lazy<DaemonWorker>(() => new DaemonWorker());

        /// <summary>
        /// tasks waiting to be executed.
        /// </summary>
        private readonly ConcurrentQueue<DaemonAction> tasks = new ConcurrentQueue<DaemonAction>();

        /// <summary>
        /// A public accessor for my instance.
        /// </summary>
        public static DaemonWorker Instance { get { return lazy.Value; } }

        /// <summary>
        /// A way for outside objects to look at the length of the queue.
        /// </summary>
        public int QueueLength => tasks.Count;

        /// <summary>
        /// The period (in milliseconds) for which I sleep between jobs.
        /// </summary>
        public readonly int IntervalMs = 5000;

        /// <summary>
        /// Construct (the one, singleton) instance of the DaemonWorker class
        /// </summary>
        private DaemonWorker() : base("Daemon", Globals.ThisAddIn.Log)
        {
            Interval = TimeSpan.FromMilliseconds(IntervalMs);
            this.Start();
        }

        /// <summary>
        /// Add a task to my queue.
        /// </summary>
        /// <param name="task">The task to add.</param>
        public void AddTask(DaemonAction task)
        {
            tasks.Enqueue(task);
        }

        /// <summary>
        /// Return an enumerable of the descriptions of my pending tasks.
        /// </summary>
        /// <returns> an enumerable of the descriptions of my pending tasks.</returns>
        public IEnumerable<string> PendingTaskDescriptions => this.tasks.Select(t => t.Description).ToList();

        /// <summary>
        /// Put me into a mode where I finish all the work I have to do quickly.
        /// </summary>
        /// <returns></returns>
        public override int PrepareShutdown()
        {
            int shutDownInterval = 10;
            if (this.Interval.Milliseconds > shutDownInterval)
            {
                this.Interval = TimeSpan.FromMilliseconds(shutDownInterval);
            }
            return this.QueueLength;
        }

        /// <summary>
        /// Take one task from the queue (if any), and perform it.
        /// </summary>
        internal override void PerformIteration()
        {
            DaemonAction task;

            if (tasks.TryDequeue(out task))
            {
                Log.Info($"About to perform {task.Description}.");

                try
                {
                    string report = task.Perform();
                    Log.Info($"{task.Description} completed: {report}");
                }
                catch (ActionRetryableException retryable)
                {
                    if (++task.Attempts < task.MaxAttempts)
                    {
                        tasks.Enqueue(task);
                        Log.Warn($"{task.Description} failed with error {retryable.GetType().Name}: {retryable.Message}; requeueing");
                    }
                    else
                    {
                        ErrorHandler.Handle($"{task.Description} failed with error {retryable.GetType().Name}: {retryable.Message}; too many retries, aborting", 
                            retryable, task.NotifyOnFailure);
                    }
                }
                catch (Exception any)
                {
                    ErrorHandler.Handle($"{task.Description} failed with error {any.GetType().Name}: {any.Message}; Not retryable, aborting", 
                        any, task.NotifyOnFailure);
                }
            }
        }
    }
}
