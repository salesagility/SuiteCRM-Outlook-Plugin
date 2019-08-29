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
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System.Runtime.InteropServices;
    using System.Collections.Generic;
    using SuiteCRMClient;
    using System.Threading;

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
    /// <see cref="Daemon.TransmitNewAction{OutlookItemType, SyncStateType}"/>.</item>
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
    /// </remarks>
    /// <typeparam name="ItemType">The type of the item to be/being synced.</typeparam>
    public abstract class SyncState<ItemType> : SyncState
        where ItemType : class
    {
        /// <summary>
        /// Underlying store for my <see cref="OutlookItem"/> property. 
        /// </summary>
        protected ItemType Item;

        /// <summary>
        /// Handle onto the MAPI namespace.
        /// </summary>
        private static Outlook.NameSpace mapiNs = null;

        /// <summary>
        /// Handle onto the MAPI namespace, guaranteed to exist.
        /// </summary>
        protected static Outlook.NameSpace MapiNs => mapiNs ?? (mapiNs = Globals.ThisAddIn.Application.GetNamespace("MAPI"));

        public abstract Outlook.Folder DefaultFolder { get; }

        /// <summary>
        /// The outlook item for which I maintain the synchronisation state.
        /// </summary>
        public ItemType OutlookItem {
            get
            {
                return this.VerifyItem() ? this.Item : null;
            }
            private set
            {
                this.Item = value;
            }
        }

        /// <summary>
        /// Varify that item has not become detached (does not throw a COMException when interrogated).
        /// </summary>
        /// <returns>false </returns>
        public abstract bool VerifyItem();

        /// <summary>
        /// A lock that should be obtained before operations which operate on the TxState or the
        /// cached value.
        /// </summary>
        private object txStateLock = new object();

        /// <summary>
        /// Create a new instance of a SyncState wrapping this item.
        /// </summary>
        /// <param name="item">The item to wrap.</param>
        /// <param name="itemId">The EntryId of that item.</param>
        /// <param name="crmId">The CRM Id of that item, if known, else null.</param>
        /// <param name="modifiedDate">When that item was last modified.</param>
        public SyncState(ItemType item, string itemId, CrmId crmId, DateTime modifiedDate): base(itemId)
        {
            this.OutlookItem = item;
            this.CrmEntryId = crmId;
            this.OModifiedDate = modifiedDate;
        }

        /// <remarks>
        /// The state transition engine. If we're building a DEBUG build, log all state transitions;
        /// OBVIOUSLY, the semantics of this, apart from the side effect, must be identical between
        /// DEBUG and non DEBUG builds.
        /// </remarks>
#if DEBUG
        private TransmissionState ts = TransmissionState.NewFromOutlook;
        internal TransmissionState TxState 
        { 
            get { return this.ts; } 
            private set {
                Globals.ThisAddIn.Log.Debug(
                    $"{this.GetType().Name} '{this.Cache?.Description}': transition {this.ts} => {value}");
                this.ts = value;
            }
        }
#else
        internal TransmissionState TxState { get; private set; } = TransmissionState.NewFromOutlook;
#endif


        /// <summary>
        /// The cache of the state of the item when it was last synced.
        /// </summary>
        public ProtoItem<ItemType> Cache { get; protected set; }


        /// <summary>
        /// A string constructed from fields which uniquely describe my item.
        /// </summary>
        public abstract string IdentifyingFields { get; }

        /// <summary>
        /// True if the Outlook item wrapped by this state may be synchronised even when synchronisation is set to none.
        /// </summary>
        /// <remarks>
        /// At present, only Contacts have the manual override mechanism.
        /// </remarks>
        public virtual bool IsManualOverride => false;

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
                var older = this.Cache.AsNameValues()
                    .AsDictionary();

                var current = this.CreateProtoItem()
                    .AsNameValues()
                    .AsDictionary();
                unchanged = current != null && older.Keys.Count.Equals(current.Keys.Count);

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
        /// Create an appropriate proto-item for my outlook item.
        /// </summary>
        /// <returns>the proto-item.</returns>
        internal abstract ProtoItem<ItemType> CreateProtoItem();

        /// <summary>
        /// True if the Outlook item I represent has been deleted.
        /// </summary>
        public override bool IsDeletedInOutlook
        {
            get
            {
                return !this.VerifyItem();
            }
        }


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
        internal virtual bool ShouldPerformSyncNow()
        {
            DateTime utcNow = DateTime.UtcNow;
            double modifiedSinceSeconds = Math.Abs((utcNow - OModifiedDate).TotalSeconds);
            ILogger log = Globals.ThisAddIn.Log;

            bool result;
            bool reallyChanged = this.ReallyChanged();
            bool isSyncable = this.ShouldSyncWithCrm;
            string prefix = $"SyncState.ShouldPerformSyncNow: {this.CrmType} {this.CrmEntryId}";

            log.Debug(reallyChanged ? $"{prefix} has changed." : $"{prefix} has not changed.");
            log.Debug(isSyncable ? $"{prefix} is syncable." : $"{ prefix} is not syncable.");
            log.Debug(IsManualOverride ? $"{prefix} is on manual override." : $"{prefix} is not on manual override.");

            lock (this.txStateLock)
            {
                /* result is set within the lock to prevent one thread capturing another thread's
                 * state change. */
                result = (IsManualOverride || (isSyncable && reallyChanged)) && this.TxState == TransmissionState.Pending && modifiedSinceSeconds > 2;
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
                    case TransmissionState.PresentAtStartup:
                    case TransmissionState.Pending:
                        this.LogAndSetTxState(TransmissionState.PresentAtStartup);
                        break;
                    default:
                        throw new BadStateTransition(this, this.TxState, TransmissionState.PresentAtStartup);
                }
            }
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to <see cref="TransmissionState.NewFromCRM"/>.
        /// </summary>
        internal void SetNewFromCRM()
        {
            lock (this.txStateLock)
            {
                switch (this.TxState)
                {
                    case TransmissionState.NewFromOutlook:
                        /* we OUGHT to get syncstates as soon as they're added, but because
                         * of asynchronous processing sometimes we don't, and in practice they
                         * may get as far as Queued. */
                    case TransmissionState.Pending:
                    case TransmissionState.Queued:
                        this.LogAndSetTxState(TransmissionState.NewFromCRM);
                        break;
                    default:
                        throw new BadStateTransition(this, this.TxState, TransmissionState.NewFromCRM);
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
                if (iSwearThatTransmissionHasFailed)
                {
                    this.LogAndSetTxState(TransmissionState.Pending);
                }
                else
                {
                    switch (this.TxState)
                    {
                        case TransmissionState.NewFromOutlook:
                        case TransmissionState.NewFromCRM:
                        case TransmissionState.PresentAtStartup:
                        /* any new item may, and often will, be set to 'Pending'. */
                        case TransmissionState.Pending:
                            /* If 'Pending', may remain 'Pending'. */
                        case TransmissionState.Queued:
                        /* If it's in the queue and is modified, it's probably best to
                         * go back to pending, because other modifications are likely */
                        case TransmissionState.Synced:
                        /* a synced item, if edited in Outlook, should be set to 'Pending'. */
                        case TransmissionState.PendingDeletion:
                            /* a pending deletion item should be treated like a 'Synced' item */
                            this.LogAndSetTxState(TransmissionState.Pending);
                            break;
                        default:
                            throw new BadStateTransition(this, this.TxState, TransmissionState.Pending);
                    }
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
                    case TransmissionState.PresentAtStartup:
                        this.LogAndSetTxState(TransmissionState.Queued);
                        break;
                    default:
                        throw new BadStateTransition(this, this.TxState, TransmissionState.Queued);
                }
            }
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to <see cref="TransmissionState.Transmitted"/>.
        /// </summary>
        internal virtual void SetTransmitted()
        {
            lock (this.txStateLock)
            {
                switch (this.TxState)
                {
                    case TransmissionState.Queued:
                        this.LogAndSetTxState(TransmissionState.Transmitted);
                        break;
                    default:
                        throw new BadStateTransition(this, this.TxState, TransmissionState.Transmitted);
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
                            this.Cache = this.CreateProtoItem();
                            this.LogAndSetTxState(TransmissionState.Synced);
                            this.OModifiedDate = DateTime.UtcNow;
                            break;
                        default:
                            throw new BadStateTransition(this, this.TxState, TransmissionState.Synced);
                    }
                }
            }
        }


        /// <summary>
        /// Set the transmission state of this SyncState object to <see cref="TransmissionState.Synced"/> 
        /// and its CRM entry ID to this crmEntryId, and recache its Outlook item.
        /// </summary>
        /// <param name="crmEntryId">The id of the object in CRM.</param>
        internal void SetSynced(CrmId crmEntryId)
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
                        throw new BadStateTransition(this, this.TxState, TransmissionState.PendingDeletion);
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
            try
            {
                if (this.Cache == null)
                {
                    this.Cache = this.CreateProtoItem();
                }
                Globals.ThisAddIn.Log.Debug(
                    $"{this.GetType().Name} '{this.Cache?.Description}': transition {this.TxState} => {newState}");
            }
            catch (InvalidComObjectException)
            {
                // ignore. It doesn't matter. Although TODO: I'd love to know what happens.
            }
            this.TxState = newState;
        }
    }
}
