using SuiteCRMClient.Logging;

namespace SuiteCRMAddIn
{
    using System.IO;

    public class FileLogger : ILogger
    {
        private readonly string _logDirPath;

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

            using (var fileStream = OpenOrCreateFileStream(logFileInfo, logFilePath))
            using (var log = new StreamWriter(fileStream))
            {
                log.WriteLine(logMessage);
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
