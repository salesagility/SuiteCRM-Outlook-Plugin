using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SuiteCRMClient.Logging;

namespace SuiteCRMAddIn.BusinessLogic
{
    /// <summary>
    /// A thing which maintains a queue of tasks and executes them in turn. Explicitly 
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
        private ConcurrentQueue<DaemonAction> tasks = new ConcurrentQueue<DaemonAction>();

        /// <summary>
        /// A public accessor for my instance.
        /// </summary>
        public static DaemonWorker Instance { get { return lazy.Value; } }

        /// <summary>
        /// Construct (the one, singleton) instance of the DaemonWorker class
        /// </summary>
        private DaemonWorker() : base("Daemon", Globals.ThisAddIn.Log)
        {
            SyncPeriod = TimeSpan.FromSeconds(30);
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

        internal override void PerformIteration()
        {
            DaemonAction task;

            if (tasks.TryDequeue(out task))
            {
                Log.Info($"About to perform {task.Description}");

                Robustness.DoOrLogError(this.Log, () => task.Perform());
            }
        }
    }
}
