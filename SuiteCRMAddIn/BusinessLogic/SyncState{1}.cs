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
    using Exceptions;
    using ProtoItems;
    using SuiteCRMClient.Logging;
    using System;
    using Microsoft.Office.Interop.Outlook;
    using System.Collections.Generic;

    /// <summary>
    /// The sync state of an item of the specified type.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Contrary to appearances this file is not a backup or a mistake but is vital to
    /// the working of the system!
    /// </para>
    /// <para>
    /// The SyncState object is essentially a handle that a Synchroniser holds on an
    /// Outlook item, on which it hangs information to track the state of synchronisation
    /// of that item.
    /// </para>
    /// <para>
    /// The life-cycle of a sync state is as follows:
    /// <list type="ordered">
    /// <item>On creation, it is in the state <see cref="TransmissionState.NewFromOutlook"/></item>
    /// <item>If a state is created from a CRM record, it is set to <see cref="TransmissionState.NewFromCRM"/></item>
    /// <item>When a change has been made to it Outlook side, the relevant 
    /// <see cref="Synchroniser{OutlookItemType, SyncStateType}"/> sets it to <see cref="TransmissionState.Pending"/> , and passes it to an
    /// <see cref="Daemon.AbstractTransmissionAction{OutlookItemType, SyncStateType}"/>.</item>
    /// <item>The AbstractTransmissionAction sets it to <see cref="TransmissionState.Queued"/> , and in due course) sends it back to 
    /// the same Synchroniser.</item>
    /// <item>The Synchroniser sets it to <see cref="TransmissionState.Transmitted"/> and transmits 
    /// it to CRM; if transmission succeeds it sets it to 
    /// <see cref="TransmissionState.Synced"/>, otherwise to <see cref="TransmissionState.Pending"/>.</item>
    /// <item>After failure of transmission, if the AbstractTransmissionAction has any 
    /// attempts left, it resets the state to <see cref="TransmissionState.Queued"/> and presently 
    /// tries again.</item>
    /// <item>Periodically, the relevant <see cref="Synchroniser{OutlookItemType, SyncStateType}"/> performs a 
    /// sweep of all items which are in <see cref="TransmissionState.Pending"/> and transmits them; 
    /// these are items which had presumably failed to be synced earlier.</item>
    /// </list>
    /// </para>
    /// <typeparam name="ItemType">The type of the item to be/being synced.</typeparam>
    public abstract class SyncState<ItemType> : SyncState
    {
        /// <summary>
        /// The outlook item for which I maintain the synchronisation state.
        /// </summary>
        public ItemType OutlookItem { get; private set; }

        /// <summary>
        /// A lock that should be obtained before operations which operate on the TxState or the
        /// cached value.
        /// </summary>
        private object txStateLock = new object();


        public SyncState(ItemType item, string crmId, DateTime modifiedDate)
        {
            this.OutlookItem = item;
            this.CrmEntryId = crmId;
            this.OModifiedDate = modifiedDate;
        }

        /// <remarks>
        /// The state transition engine.
        /// </remarks>
        internal TransmissionState TxState { get; private set; } = TransmissionState.NewFromOutlook;


        /// <summary>
        /// The cache of the state of the item when it was last synced.
        /// </summary>
        public ProtoItem<ItemType> Cache { get; protected set; }

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
        /// <remarks>Beware inverted logic here.</remarks>
        /// <returns>True if the object has changed.</returns>
        internal virtual bool ReallyChanged()
        {
            bool unchanged;

            if (this.Cache == null)
            {
                unchanged = false;
            }
            else
            {
                var older = this.Cache.AsNameValues(this.CrmEntryId)
                    .AsDictionary();

                var current = this.CreateProtoItem(this.OutlookItem)
                    .AsNameValues(this.CrmEntryId)
                    .AsDictionary();
                unchanged = older.Keys.Count.Equals(current.Keys.Count);

                if (unchanged)
                {
                    foreach (string key in older.Keys)
                    {
                        unchanged &= current.ContainsKey(key) &&
                            ((older[key] == null && current[key] == null) ||
                             (older[key] != null && older[key].Equals(current[key])));
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
            DateTime utcNow = DateTime.UtcNow;
            double modifiedSinceSeconds = Math.Abs((utcNow - OModifiedDate).TotalSeconds);
            ILogger log = Globals.ThisAddIn.Log;
            bool reallyChanged = this.ReallyChanged();
            bool isSyncable = this.ShouldSyncWithCrm;
            string prefix = $"SyncState.ShouldPerformSyncNow: {this.CrmType} {this.CrmEntryId}";

            log.Debug(reallyChanged ? $"{prefix} has changed." : $"{prefix} has not changed.");
            log.Debug(isSyncable ? $"{prefix} is syncable." : $"{ prefix} is not syncable.");

            bool result;

            lock (this.txStateLock)
            {
                /* result is set within the lock to prevent one thread capturing another thread's
                 * state change. */
                result = isSyncable && reallyChanged && this.TxState == TransmissionState.Pending && modifiedSinceSeconds > 2;
                if (result)
                {
                    this.OModifiedDate = utcNow;
                }
            }

            log.Debug(this.TxState == TransmissionState.Pending ? $"{prefix} is recently updated" : $"{prefix} is not recently updated");

            log.Debug(result ? $"{prefix} should be synced now" : $"{prefix} should not be synced now");

            return result;
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to <see cref="TransmissionState.PresentAtStartup"/>.
        /// </summary>
        internal void SetPresentAtStartup()
        {
            lock (this.txStateLock)
            {
                switch (this.TxState)
                {
                    case TransmissionState.NewFromOutlook:
                    case TransmissionState.Pending:
                        this.LogAndSetTxState(TransmissionState.PresentAtStartup);
                        break;
                    default:
                        throw new BadStateTransition($"{this.TxState} => PresentAtStartup");
                }
            }
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to <see cref="TransmissionState.PresentAtStartup"/>.
        /// </summary>
        internal void SetNewFromCRM()
        {
            lock (this.txStateLock)
            {
                switch (this.TxState)
                {
                    case TransmissionState.NewFromOutlook:
                    case TransmissionState.Pending:
                        this.LogAndSetTxState(TransmissionState.NewFromOutlook);
                        break;
                    default:
                        throw new BadStateTransition($"{this.TxState} => NewFromCRM");
                }
            }
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to <see cref="TransmissionState.Pending"/>.
        /// </summary>
        /// <param name="iSwearThatTransmissionHasFailed">Allows override of state transition 
        /// flow ONLY when transmission has failed.</param>
        internal void SetPending(bool iSwearThatTransmissionHasFailed = false)
        {
            lock (this.txStateLock)
            {
                if (iSwearThatTransmissionHasFailed && this.TxState == TransmissionState.Transmitted)
                {
                    this.LogAndSetTxState(TransmissionState.Pending);
                }
                switch (this.TxState)
                {
                    case TransmissionState.NewFromOutlook:
                    case TransmissionState.PresentAtStartup:
                        /* a new item may, and often will, be set to 'Pending'. */
                    case TransmissionState.Pending:
                        /* If 'Pending', may remain 'Pending'. */
                    case TransmissionState.Synced:
                        /* a synced item, if edited in Outlook, should be set to 'Pending'. */
                    case TransmissionState.PendingDeletion:
                        /* a pending deletion item should be treated like a 'Synced' item */
                        this.LogAndSetTxState(TransmissionState.Pending);
                        break;
                    default:
                        throw new BadStateTransition($"{this.TxState} => Pending");
                }
            }
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to <see cref="TransmissionState.Queued"/>.
        /// </summary>
        internal void SetQueued()
        {
            lock (this.txStateLock)
            {
                switch (this.TxState)
                {
                    case TransmissionState.Pending:
                        this.LogAndSetTxState(TransmissionState.Queued);
                        break;
                    default:
                        throw new BadStateTransition($"{this.TxState} => Queued");
                }
            }
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to <see cref="TransmissionState.Transmitted"/>.
        /// </summary>
        internal void SetTransmitted()
        {
            lock (this.txStateLock)
            {
                switch (this.TxState)
                {
                    case TransmissionState.Queued:
                        this.LogAndSetTxState(TransmissionState.Transmitted);
                        break;
                    default:
                        throw new BadStateTransition($"{this.TxState} => Transmitted");
                }
            }
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to <see cref="TransmissionState.Synced"/>, 
        /// and recache its Outlook item.
        /// </summary>
        /// <param name="iSwearReceivedFromCRM">When an item is received from CRM, it is created or updated in 
        /// Outlook and is likely to be queued for retransmission back to CRM before SetSynced() happens. 
        /// Under this situation ONLY, it's allowable to move from any to Synced.</param>
        internal void SetSynced(bool iSwearReceivedFromCRM = false)
        {
            lock (this.txStateLock)
            {
                if (iSwearReceivedFromCRM)
                {
                    this.TxState = TransmissionState.Synced;
                }
                else
                {
                    switch (this.TxState)
                    {
                        case TransmissionState.NewFromCRM:
                            /* if ol item is created from a CRM record, it will be 'NewFromCRM' then 'Synced' */
                        case TransmissionState.PresentAtStartup:
                            /* when the add-in first starts up, new SyncStates will get synced. */
                        case TransmissionState.Synced:
                            /* if ol item is unchanged but CRM record is changed, it will be 'Synced' then 'Synced' */
                        case TransmissionState.Transmitted:
                            /* if a state has been marked as pending deletion and then is found on the next 
                             * synchronisation run, it should be set back to synced */
                        case TransmissionState.PendingDeletion:
                            /* if ol item is transmitted to CRM, it will be 'Transmitted' then 'Synced' */
                            this.Cache = this.CreateProtoItem(this.OutlookItem);
                            this.LogAndSetTxState(TransmissionState.Synced);
                            this.OModifiedDate = DateTime.UtcNow;
                            break;
                        default:
                            throw new BadStateTransition($"{this.TxState} => Synced");
                    }
                }
            }
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to <see cref="TransmissionState.Synced"/> 
        /// and its CRM entry ID to this crmEntryId, and recache its Outlook item.
        /// </summary>
        /// <param name="crmEntryId">The id of the object in CRM.</param>
        internal void SetSynced(string crmEntryId)
        {
            this.SetSynced();
            this.CrmEntryId = crmEntryId;
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to <see cref="TransmissionState.PendingDeletion"/>.
        /// </summary>
        internal void SetPendingDeletion()
        { 
            lock (this.txStateLock)
            {
                switch (this.TxState)
                {
                    case TransmissionState.NewFromOutlook:
                        /* if a CRM outlook is deleted while Outlook is offline,
                         * you'll get this sequence. */
                    case TransmissionState.Synced:
                        this.LogAndSetTxState(TransmissionState.PendingDeletion);
                        break;
                    default:
                        throw new BadStateTransition($"{this.TxState} => PendingDeletion");
                }
            }
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to <see cref="TransmissionState.Invalid"/>.
        /// </summary>
        /// <remarks>
        /// Removes the invalid sync state from the caches, which will hopefully allow
        /// the system to stabilise itself.
        /// </remarks>
        internal void SetInvalid()
        {
            this.TxState = TransmissionState.Invalid;
            // SyncStateFactory.Instance.RemoveSyncState(this);
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to pending, if it has really changed.
        /// </summary>
        internal void SetPendingIfReallyChanged()
        {
            if (this.ReallyChanged())
            {
                Robustness.DoOrLogError(
                    Globals.ThisAddIn.Log,
                    () => this.SetPending(),
                    $"{this.GetType().Name}.SetPendingIfReallyChanged");
            }
        }

        /// <summary>
        /// We should not copy changes from CRM to Outlook if changes from Outlook to CRM are queued.
        /// </summary>
        /// <returns>True if state is synced.</returns>
        internal bool MayBeUpdatedFromCRM()
        {
            bool result;

            switch (this.TxState)
            {
                case TransmissionState.NewFromOutlook:
                case TransmissionState.Synced:
                    result = true;
                    break;
                default:
                    result = false;
                    break;
            }

            return result;
        }

        private void LogAndSetTxState(TransmissionState newState)
        {
#if DEBUG
            Globals.ThisAddIn.Log.Debug($"{this.Cache?.Description}: transition {this.TxState} => {newState}");
#endif
            this.TxState = newState;
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
            NewFromOutlook,
            /// <summary>
            /// This is a SyncState representing an outlook item which was present when 
            /// Outlook was started.
            /// </summary>
            PresentAtStartup,
            /// <summary>
            /// This is a SyncState object representing a outlook item which has just been 
            /// created from a CRM item.
            /// </summary>
            NewFromCRM,
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
            Synced,
            /// <summary>
            /// A state is put into state PendingDeletion if it is not found in CRM at 
            /// one synchronisation run; if it is not found in the subsequent run and is
            /// still in state PendingDeletion, then it should be deleted.
            /// </summary>
            PendingDeletion,
            /// <summary>
            /// The sync state is in an invalid state and should never be synced.
            /// </summary>
            Invalid
        }
    }
}
