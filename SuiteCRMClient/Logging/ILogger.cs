using System;

namespace SuiteCRMClient.Logging
{
    /// <summary>
    /// Interface implemented by all logging classes
    /// 
    /// </summary>
    public interface ILogger : IDisposable
    {
        /// <summary>
        /// Adds the entry.
        /// 
        /// </summary>
        /// <param name="logMessage">The log message.</param><param name="type">The type.</param>
        void AddEntry(string logMessage, LogEntryType type);

        /// <summary>
        /// Expose the logging level.
        /// </summary>
        LogEntryType Level
        {
            get; set;
        }
    }
}