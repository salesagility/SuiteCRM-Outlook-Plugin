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
    using System;
    using log4net;
    using log4net.Appender;
    using log4net.Core;
    using log4net.Layout;
    using log4net.Repository.Hierarchy;
    using SuiteCRMClient.Logging;
    using System.Text;
    using System.Collections.Generic;
    using log4net.Repository;

    public class Log4NetLogger: SuiteCRMClient.Logging.AbstractLogger
    {
        private readonly ILog log;

        private Log4NetLogger(string area)
        {
            log = LogManager.GetLogger(area);
        }

        /// <summary>
        /// Expose the logging level.
        /// </summary>
        public override LogEntryType Level
        {
            get
            {
                Logger lumberjack = (Logger)this.log.Logger;
                return ToLogEntryType(lumberjack.Level);
            }
            set
            {
                Logger lumberjack = (Logger)this.log.Logger;
                lumberjack.Level = FromLogEntryType(value);
            }
        }

        /// <summary>
        /// Translate a log4net level to a LogEntryType.
        /// </summary>
        /// <param name="level">The level.</param>
        /// <returns>The corresponding entry type.</returns>
        private static LogEntryType ToLogEntryType(Level level)
        {
            LogEntryType result;

            if (level.CompareTo(log4net.Core.Level.Debug) == 0)
            {
                result = LogEntryType.Debug;
            }
            else if (level.CompareTo(log4net.Core.Level.Info) < 0)
            {
                result = LogEntryType.Information;
            }
            else if (level.CompareTo(log4net.Core.Level.Warn) < 0)
            {
                result = LogEntryType.Warning;
            }
            else 
            {
                result = LogEntryType.Error;
            }

            return result;
        }

        /// <summary>
        /// Convert a LogEntryType to the corresponding log4net level.
        /// </summary>
        /// <param name="entryType">An entry type.</param>
        /// <returns>the corresponding log4net level</returns>
        private static Level FromLogEntryType(LogEntryType entryType)
        {
            Level result;

            switch (entryType)
            {
                case LogEntryType.Debug:
                    result = log4net.Core.Level.Debug;
                    break;
                case LogEntryType.Information:
                    result = log4net.Core.Level.Info;
                    break;
                case LogEntryType.Warning:
                    result = log4net.Core.Level.Warn;
                    break;
                default:
                    result = log4net.Core.Level.Error;
                    break;
            }

            return result;
        }

        public static Log4NetLogger FromFilePath(string area, string filePath, Func<IEnumerable<string>> headerFunction, LogEntryType entryType)
        {
            var hierarchy = (Hierarchy)LogManager.GetRepository();

            var patternLayout = new PatternLayoutWithHeader("%date | %-2thread | %-5level | %message%newline", headerFunction);
            patternLayout.ActivateOptions();

            var level = FromLogEntryType(entryType);
            var appender = new RollingFileAppender
            {
                AppendToFile = true,
                File = filePath,
                Layout = patternLayout,
                RollingStyle = RollingFileAppender.RollingMode.Size,
                MaxFileSize = 1000000, // 1MB
                StaticLogFileName = true,
                MaxSizeRollBackups = 10,
                Threshold = level,
                Encoding = Encoding.UTF8,
            };
            appender.ActivateOptions();

            hierarchy.Root.AddAppender(appender);
            hierarchy.Root.Level = level;
            hierarchy.Configured = true;

            return new Log4NetLogger(area);
        }

        public override void AddEntry(string message, LogEntryType type)
        {
#if DEBUG
            log.Debug($"Current memory usage: {System.Diagnostics.Process.GetCurrentProcess().WorkingSet64/1000000} Mb");
#endif

            switch (type)
            {
                case LogEntryType.Debug:
                    log.Debug(message);
                    return;

                case LogEntryType.Error:
                    log.Error(message);
                    return;

                case LogEntryType.Information:
                    log.Info(message);
                    return;

                case LogEntryType.Warning:
                    log.Warn(message);
                    return;
            }
        }

        /// <summary>
        /// Flush all buffers.
        /// </summary>
        /// <remarks>
        /// Thanks to http://stackoverflow.com/questions/2045935/is-there-anyway-to-programmably-flush-the-buffer-in-log4net
        /// </remarks>
        public void FlushBuffers()
        {
            ILoggerRepository rep = LogManager.GetRepository();
            foreach (IAppender appender in rep.GetAppenders())
            {
                var buffered = appender as BufferingAppenderSkeleton;
                if (buffered != null)
                {
                    buffered.Flush();
                }
            }
        }

        /// <summary>
        /// Make sure the last items logged get output.
        /// </summary>
        public override void Dispose()
        {
            this.FlushBuffers();
        }

        private class PatternLayoutWithHeader : PatternLayout
        {
            private readonly Func<IEnumerable<string>> _headerFunction;

            public PatternLayoutWithHeader(string pattern, Func<IEnumerable<string>> headerFunc)
                : base(pattern)
            {
                _headerFunction = headerFunc;
            }

            public override string Header
            {
                get
                {
                    const string separator = "-----------------------------";
                    var newline = Environment.NewLine;
                    return
                        separator + newline +
                        string.Join(newline, _headerFunction()) + newline +
                        separator + newline;
                }
            }
        }
    }
}
