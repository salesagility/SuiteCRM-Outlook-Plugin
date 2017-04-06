
using Microsoft.Office.Interop.Outlook;
using SuiteCRMAddIn.ProtoItems;
using SuiteCRMClient.Logging;
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
    /// The sync state of an item of the specified type. Contrary to appearances this 
    /// file is not a backup or a mistake but is vital to the working of the system!
    /// </summary>
    /// <typeparam name="ItemType">The type of the item to be/being synced.</typeparam>
    public abstract class SyncState<ItemType> : SyncState
    {
        /// <summary>
        /// Backing store for the OutlookItem property.
        /// </summary>
        private ItemType olItem;

        /// <summary>
        /// The outlook item for which I maintain the synchronisation state.
        /// </summary>
        public ItemType OutlookItem
        {
            get
            {
                return olItem;
            }
            set
            {
                olItem = value;
                this.Cache = this.CreateProtoItem(value);
            }
        }

        /// <summary>
        /// The cache of the state of the item when it was first linked.
        /// </summary>
        public ProtoItem<ItemType> Cache { get; private set; }

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
        protected virtual bool ReallyChanged()
        {
            var older = this.Cache.AsNameValues(this.OutlookItemEntryId)
                .AsDictionary();
            var current = this.CreateProtoItem(this.OutlookItem)
                    .AsNameValues(this.OutlookItemEntryId)
                    .AsDictionary();
            bool result = older.Keys.Count.Equals(current.Keys.Count);

            foreach (string key in older.Keys){
                result &= (older[key] == null && current[key] == null) ||
                    older[key].Equals(current[key]);
            }

            return result;
        }

        /// <summary>
        /// Create an appropriate proto-item for this outlook item.
        /// </summary>
        /// <param name="outlookItem">The outlook item to copy.</param>
        /// <returns>the proto-item.</returns>
        internal abstract ProtoItem<ItemType> CreateProtoItem(ItemType outlookItem);

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
            DateTime utcNow = DateTime.UtcNow;
            double modifiedSinceSeconds = Math.Abs((utcNow - OModifiedDate).TotalSeconds);
            ILogger log = Globals.ThisAddIn.Log;
            bool reallyChanged = this.ReallyChanged();
            bool shouldSync = this.ShouldSyncWithCrm;
            string prefix = $"SyncState.ShouldPerformSyncNow: {this.CrmType} {this.CrmEntryId}";

            log.Debug(reallyChanged ? $"{prefix} has changed." : $"{prefix} has not changed.");
            log.Debug(shouldSync ? $"{prefix} should be synced." : $"{ prefix} shouldSync not be synced.");

            if (modifiedSinceSeconds > 5 || modifiedSinceSeconds > 2 && this.IsUpdate == 0)
            {
                this.OModifiedDate = utcNow;
                this.IsUpdate = 1;
            }

            log.Debug(IsUpdate == 1 ? $"{prefix} is recently updated" : $"{prefix} is not recently updated");

            var result = this.IsUpdate == 1 && shouldSync && reallyChanged;

            log.Debug(result ? $"{prefix} should be synced now" : $"{prefix} should not be synced now");

            return result;
        }
    }
}
