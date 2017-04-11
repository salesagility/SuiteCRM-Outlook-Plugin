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
    using SuiteCRMAddIn.ProtoItems;
    using SuiteCRMClient.Logging;
    using System;

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
        /// A lock that should be obtained before operations which operate on the TxState or the
        /// cached value.
        /// </summary>
        private object txStateLock = new object();

        /// <remarks>
        /// Legacy code. Neither Andrew Forrest nor I (Simon Brooke) really understand what this 
        /// is about; its values are small integers and probably ought to be an enum, but we don't 
        /// know what the values mean.
        /// </remarks>
        internal TransmissionState TxState { get; private set; } = TransmissionState.New;

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
        /// <returns>True if the object has changed.</returns>
        protected virtual bool ReallyChanged()
        {
            bool unchanged;

            if (this.Cache == null)
            {
                unchanged = false;
            }
            else
            {
                var older = this.Cache.AsNameValues(this.OutlookItemEntryId)
                    .AsDictionary();
                var current = this.CreateProtoItem(this.OutlookItem)
                        .AsNameValues(this.OutlookItemEntryId)
                        .AsDictionary();
                unchanged = older.Keys.Count.Equals(current.Keys.Count);

                if (unchanged)
                {
                    foreach (string key in older.Keys)
                    {
                        unchanged &= (older[key] == null && current[key] == null) ||
                            older[key].Equals(current[key]);
                    }
                }
            }

            return !unchanged;
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
            bool result;
            DateTime utcNow = DateTime.UtcNow;
            double modifiedSinceSeconds = Math.Abs((utcNow - OModifiedDate).TotalSeconds);
            ILogger log = Globals.ThisAddIn.Log;
            bool reallyChanged = this.ReallyChanged();
            bool shouldSync = this.ShouldSyncWithCrm;
            string prefix = $"SyncState.ShouldPerformSyncNow: {this.CrmType} {this.CrmEntryId}";

            log.Debug(reallyChanged ? $"{prefix} has changed." : $"{prefix} has not changed.");
            log.Debug(shouldSync ? $"{prefix} should be synced." : $"{ prefix} should not be synced.");

            lock (this.txStateLock)
            {
                if (modifiedSinceSeconds > 5 || modifiedSinceSeconds > 2 &&
                (this.TxState == TransmissionState.New || this.TxState == TransmissionState.Synced))
                {
                    this.OModifiedDate = utcNow;
                    this.TxState = TransmissionState.Pending;
                }

                /* result is set within the lock to prevent one thread capturing another thread's
                 * state change. */
                result = this.TxState == TransmissionState.Pending && shouldSync && reallyChanged;
            }

            log.Debug(this.TxState == TransmissionState.Pending ? $"{prefix} is recently updated" : $"{prefix} is not recently updated");

            log.Debug(result ? $"{prefix} should be synced now" : $"{prefix} should not be synced now");

            return result;
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to pending.
        /// </summary>
        internal void SetPending()
        {
            lock (this.txStateLock)
            {
                this.TxState = TransmissionState.Pending;
            }
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to queued.
        /// </summary>
        internal void SetQueued()
        {
            lock (this.txStateLock)
            {
                this.TxState = TransmissionState.Queued;
            }
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to synced, and recache its Outlook item.
        /// </summary>
        internal void SetSynced()
        {
            lock (this.txStateLock)
            {
                this.Cache = this.CreateProtoItem(this.OutlookItem);
                this.TxState = TransmissionState.Synced;
                this.OModifiedDate = DateTime.UtcNow;
            }
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to synced and its CRM entry ID to this
        /// crmEntryId, and recache its Outlook item.
        /// </summary>
        /// <param name="crmEntryId">The id of the object in CRM.</param>
        internal void SetSynced(string crmEntryId)
        {
            this.SetSynced();
            this.CrmEntryId = crmEntryId;
        }

        /// <summary>
        /// Set the transmission state of this SyncState object to transmitted.
        /// </summary>
        internal void SetTransmitted()
        {
            lock (this.txStateLock)
            {
                this.TxState = TransmissionState.Transmitted;
            }
        }


        /// <summary>
        /// States a SyncState object can be in with regard to transmission and synchronisation
        /// with CRM. See TxState.
        /// </summary>
        public enum TransmissionState
        {
            /// <summary>
            /// This is a new SyncState object which has not yet been transmitted.
            /// </summary>
            New,
            /// <summary>
            /// A change has been registered on this SyncState object but it has 
            /// not been transmitted.
            /// </summary>
            Pending,
            /// <summary>
            /// This SyncState has been queued for transmission but has not yet been 
            /// transmitted.
            /// </summary>
            Queued,
            /// <summary>
            /// The Outlook item associated with this SyncState has been transmitted,
            /// but no confirmation has yet been received that it has been accepted.
            /// </summary>
            Transmitted,
            /// <summary>
            /// The Outlook item associated with this SyncState has been transmitted
            /// and accepted by CRM.
            /// </summary>
            Synced
        }
    }
}
