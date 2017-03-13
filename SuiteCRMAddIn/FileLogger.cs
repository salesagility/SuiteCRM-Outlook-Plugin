using SuiteCRMClient.Logging;

namespace SuiteCRMAddIn
{
    using System;
    using System.IO;

    public class FileLogger : ILogger
    {
        private readonly string _logDirPath;

        public LogEntryType level = LogEntryType.Error;

        public LogEntryType Level
        {
            get
            {
                return this.level;
            }

            set
            {
                this.level = value;
            }
        }

        public FileLogger(string logDirPath)
        {
            _logDirPath = logDirPath;
        }

        public void AddEntry(string logMessage, LogEntryType type)
        {
            var logFilePath = _logDirPath + "Log-" + System.DateTime.Today.ToString("MM-dd-yyyy") + "." + "txt";
            var logFileInfo = new FileInfo(logFilePath);
            var logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
            if (!logDirInfo.Exists) logDirInfo.Create();

            using (var fileStream = OpenOrCreateFileStream(logFileInfo, logFilePath)) {
                using (var log = new StreamWriter(fileStream))
                {
                    switch (type)
                    {
                        case LogEntryType.Debug:
                            if (this.level == LogEntryType.Debug)
                            {
                                log.WriteLine(logMessage);
                            }
                            break;
                        case LogEntryType.Information:
                            if (this.level == LogEntryType.Debug ||
                                this.level == LogEntryType.Information)
                            {
                                log.WriteLine(logMessage);
                            }
                            break;
                        case LogEntryType.Warning:
                            if (this.level == LogEntryType.Debug ||
                                this.level == LogEntryType.Information ||
                                this.level == LogEntryType.Warning)
                            {
                                log.WriteLine(logMessage);
                            }
                            break;
                        default:
                            log.WriteLine(logMessage);
                            break;
                    }
                }
            }
        }

        private static FileStream OpenOrCreateFileStream(FileInfo logFileInfo, string logFilePath)
        {
            return !logFileInfo.Exists
                ? logFileInfo.Create()
                : new FileStream(logFilePath, FileMode.Append);
        }

        public void Dispose()
        {
            // Do nothing.
        }
    }
}
