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
        /// Adds this message to the log.
        /// </summary>
        /// <param name="message">The log message to add.</param>
        /// <param name="type">The type (priority) of the message.</param>
        void AddEntry(string message, LogEntryType type);

        /// <summary>
        /// Adds this message to the log, and displays it in a dialog box.
        /// </summary>
        /// <param name="message">The log message to add.</param>
        /// <param name="type">The type (priority) of the message.</param>
        void ShowAndAddEntry(string message, LogEntryType type);

        /// <summary>
        /// Expose the logging level.
        /// </summary>
        LogEntryType Level
        {
            get; set;
        }
    }
}