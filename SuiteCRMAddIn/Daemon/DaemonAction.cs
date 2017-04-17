﻿/**
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
    /// <remarks>Not currently used. You can't make a list of Outlook items detached from their 
    /// Outlook collection because they're not real objects, and if the current selection changes 
    /// before the process runs the process acts on the wrong things. I like the idea of 
    /// asynchronous processing to speed up perceived user interface response, but this isn't 
    /// working yet.</remarks>
    public interface DaemonAction
    {
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
        void Perform();
    }
}
