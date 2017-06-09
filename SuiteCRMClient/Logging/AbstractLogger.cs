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
namespace SuiteCRMClient.Logging
{
    using System.Windows.Forms;

    /// <summary>
    /// Provides a common implementation of ShowAndAddEntry for implementations of ILogger.
    /// </summary>
    public abstract class AbstractLogger : ILogger
    {
        public abstract LogEntryType Level { get; set; }

        public abstract void AddEntry(string message, LogEntryType type);
        public abstract void Dispose();

        public void ShowAndAddEntry(string message, LogEntryType type)
        {
            this.AddEntry(message, type);
            MessageBox.Show(message, type.ToString(), MessageBoxButtons.OK, IconForLogLevel(type));
        }

        /// <summary>
        /// Return an appropriate icon for this log entry type.
        /// </summary>
        /// <param name="level">The log entry type.</param>
        /// <returns>An appropriate icon.</returns>
        private static MessageBoxIcon IconForLogLevel(LogEntryType level)
        {
            MessageBoxIcon icon;

            switch (level)
            {
                case LogEntryType.Debug:
                case LogEntryType.Information:
                    icon = MessageBoxIcon.Information;
                    break;
                case LogEntryType.Warning:
                    icon = MessageBoxIcon.Warning;
                    break;
                default:
                    icon = MessageBoxIcon.Error;
                    break;
            }

            return icon;
        }
    }
}
