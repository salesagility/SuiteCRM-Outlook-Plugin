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
    using System.IO;
    using SuiteCRMClient.Logging;

    /// <summary>
    /// A logger which logs to file?
    /// </summary>
    /// <remarks>
    /// TODO: This class does not appear to be used and should probably be deleted.
    /// </remarks>
    public class FileLogger : AbstractLogger
    {
        private readonly string _logDirPath;

        public LogEntryType level = LogEntryType.Error;

        public override LogEntryType Level
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

        public override void AddEntry(string logMessage, LogEntryType type)
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

        public override void Dispose()
        {
            // Do nothing.
        }
    }
}
