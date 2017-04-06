
using Microsoft.Office.Interop.Outlook;
using System;
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
namespace SuiteCRMAddIn.BusinessLogic
{
    /// <summary>
    /// The sync state of an item specified type. Contrary to appearances this 
    /// file is not a backup or a mistake but is vital to the working of the system!
    /// </summary>
    /// <typeparam name="ItemType">The type of the item to be/being synced.</typeparam>
    public abstract class SyncState<ItemType>: SyncState
    {
        public ItemType OutlookItem { get; set; }

        /// <summary>
        /// Delete the Outlook item associated with this SyncState.
        /// </summary>
        public abstract void DeleteItem();

        /// <summary>
        /// Return true if 
        /// <list type="ordered">
        /// <item>We don't have a cached version of the related CRM item, or</item>
        /// <item>The outlook item is different from our cached version.</item>
        /// </list> 
        /// </summary>
        /// <returns></returns>
        internal abstract bool ReallyChanged();

        /// <summary>
        /// Don't send updates immediately on change, to prevent jitter; don't send updates if nothing
        /// has really changed.
        /// </summary>
        /// <remarks>
        /// The timing logic here is legacy, and neither Andrew Forrest nor I (Simon Brooke) 
        /// understand what it's intended to do; but although we've refactored it, we've left it in.
        /// </remarks>
        /// <returns>True if this item should be synced with CRM, there has been a real change, 
        /// and some time has elapsed.</returns>
        internal bool ShouldPerformSyncNow()
        {
            var utcNow = DateTime.UtcNow;
            var modifiedSinceSeconds = Math.Abs((utcNow - OModifiedDate).TotalSeconds);
            if (modifiedSinceSeconds > 5 || modifiedSinceSeconds > 2 && this.IsUpdate == 0)
            {
                this.OModifiedDate = utcNow;
                this.IsUpdate = 1;
            }

            return this.IsUpdate == 1 && this.ShouldSyncWithCrm && this.ReallyChanged();
        }

    }
}
