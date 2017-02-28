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

        public static Log4NetLogger FromFilePath(string area, string filePath, Func<IEnumerable<string>> headerFunction)
        {
            var hierarchy = (Hierarchy)LogManager.GetRepository();

            var patternLayout = new PatternLayoutWithHeader("%date | %-2thread | %-5level | %message%newline", headerFunction);
            patternLayout.ActivateOptions();

            var appender = new RollingFileAppender
            {
                AppendToFile = true,
                File = filePath,
                Layout = patternLayout,
                RollingStyle = RollingFileAppender.RollingMode.Size,
                MaxFileSize = 1000000, // 1MB
                StaticLogFileName = true,
                MaxSizeRollBackups = 10,
                Threshold = Level.Debug,
                Encoding = Encoding.UTF8,
            };
            appender.ActivateOptions();

            hierarchy.Root.AddAppender(appender);
            hierarchy.Root.Level = Level.Debug;
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
