using System;

namespace SuiteCRMClient.Logging
{
    public static class LoggerExtensions
    {
        public static void Debug(this ILogger log, string message)
        {
            log?.SafeLog(LogEntryType.Debug, message);
        }

        public static void Info(this ILogger log, string message)
        {
            log?.SafeLog(LogEntryType.Information, message);
        }

        public static void Warn(this ILogger log, string message)
        {
            log?.SafeLog(LogEntryType.Warning, message);
        }

        public static void Warn(this ILogger log, string message, Exception error)
        {
            log?.SafeLog(LogEntryType.Warning, message + "\n" + error.ToString());
        }

        public static void Error(this ILogger log, string message)
        {
            log?.SafeLog(LogEntryType.Error, message);
        }

        public static void Error(this ILogger log, string message, Exception error)
        {
            if (error == null)
            {
                Error(log, message);
            }
            else
            {
                log?.SafeLog(
                    LogEntryType.Error,
                    message + "\n" +
                    error.ToString() + "\n" +
                    "Data:" + error.Data + "\n" +
                    "HResult:" + error.HResult);
            }
        }

        private static void SafeLog(this ILogger log, LogEntryType type, string message)
        {
            try
            {
                log.AddEntry(message, type);
            }
            catch
            {
                // If we can't log, we're a bit screwed.
            }
        }
    }
}
