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

    public class Log4NetLogger: SuiteCRMClient.Logging.ILogger
    {
        private readonly ILog log;

        private Log4NetLogger(string area)
        {
            log = LogManager.GetLogger(area);
        }

        /// <summary>
        /// Expose the logging level.
        /// </summary>
        public LogEntryType Level
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

        public static Log4NetLogger FromFilePath(string area, string filePath, Func<IEnumerable<string>> headerFunction)
        {
            var hierarchy = (Hierarchy)LogManager.GetRepository();

            var patternLayout = new PatternLayoutWithHeader("%date | %-2thread | %-5level | %message%newline", headerFunction);
            patternLayout.ActivateOptions();

            var level = FromLogEntryType(Globals.ThisAddIn.Settings.LogLevel);
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

        public void AddEntry(string message, LogEntryType type)
        {
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

        public void Dispose()
        {
            // Do nothing.
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
