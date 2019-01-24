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

using System.Windows.Forms;

namespace SuiteCRMAddIn.Daemon
{
    /// <summary>
    /// The Attempts/MaxAttempts plumbing for implementing a DaemonAction.
    /// Of course, you do not need to specialise this class, but it helps.
    /// </summary>
    public abstract class AbstractDaemonAction : DaemonAction
    {
        /// <summary>
        /// If true, in the event of total failure, notify the user.
        /// </summary>
        public bool NotifyOnFailure { get; protected set; } = false;

        protected AbstractDaemonAction(int maxAttempts)
        {
            this.MaxAttempts = maxAttempts;
        }

        /// <summary>
        /// The number of times this item has been attempted.
        /// </summary>
        public int Attempts {get; set;} = 0;

        /// <summary>
        /// Get a description of this action.
        /// </summary>
        public virtual string Description => this.GetType().Name;

        /// <summary>
        /// The maximum number of times this action can be attempted before
        /// being abandoned.
        /// </summary>

        public int MaxAttempts { get; }

        /// <summary>
        /// Perform this action.
        /// </summary>
        /// <returns>A string which may be logged to report what has been done.</returns>
        /// <exception cref="System.Exception">if the performance fails.</exception>
        public abstract string Perform();
    }
}
