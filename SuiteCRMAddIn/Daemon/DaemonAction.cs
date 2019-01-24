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
namespace SuiteCRMAddIn.Daemon
{
    /// <summary>An action to be queued and performed by the DaemonWorker.</summary>
	/// <remarks>
    /// DaemonActions are intended to be run essentially once, but may be allowed a number of attempts 
	/// (intended to be limited) in case, for example due to network problems, the first attempt(s) fail. 
	/// However, DaemonActions are not intended for things which are to be run repeatedly. For that, 
	/// specialise [RepeatingProcess](class_suite_c_r_m_add_in_1_1_business_logic_1_1_repeating_process.html).
    /// </remarks>
    public interface DaemonAction
    {
        /// <summary>
        /// If true, in the event of total failure, notify the user.
        /// </summary>
        bool NotifyOnFailure { get; }

        /// <summary>
        /// The number of times this item has been attempted.
        /// </summary>
        int Attempts { get; set; }

        /// <summary>
        /// Get a description of this action.
        /// </summary>
        string Description {
            get;
        }

        /// <summary>
        /// The maximum number of times this action can be attempted before
        /// being abandoned.
        /// </summary>
        int MaxAttempts { get; }

        /// <summary>
        /// Perform this action.
        /// </summary>
        /// <returns>A string which may be logged to report what has been done.</returns>
        string Perform();
    }
}
