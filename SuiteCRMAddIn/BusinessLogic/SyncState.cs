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
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using System;
    using System.Runtime.InteropServices;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Superclass for <see cref="SyncState{ItemType}"/>, q.v..
    /// </summary>
    public abstract class SyncState : WithRemovableSynchronisationProperties
    {
        protected static Outlook.NameSpace MapiNS => Globals.ThisAddIn.Application.GetNamespace("MAPI");

        private bool _wasDeleted = false;

        public CrmId CrmEntryId { get; set; }

        public abstract string CrmType { get; }

        /// <summary>
        /// A description of the item, suitable for use in debugging logs.
        /// </summary>
        public abstract string Description { get; }

        public bool ExistedInCrm => CrmId.IsValid(CrmEntryId);

        public bool IsPublic => OutlookItemSensitivity == Outlook.OlSensitivity.olNormal;

        public DateTime OModifiedDate { get; set; }

        /// <summary>
        /// The EntryId of the Outlook item I wrap.
        /// </summary>
        public readonly string OutlookItemEntryId;

        /// <summary>
        /// True if the Outlook item I represent has been deleted.
        /// </summary>
        public abstract bool IsDeletedInOutlook { get; }

        /// <summary>
        /// Precisely 'this.OutlookItem.Sensitivity'.
        /// </summary>
        /// <remarks>Outlook item classes do not inherit from a common base class, so generic client code cannot refer to 'OutlookItem.Sensitivity'.</remarks>
        public abstract Outlook.OlSensitivity OutlookItemSensitivity { get; }

        public abstract Outlook.UserProperties OutlookUserProperties { get; }

        public virtual bool ShouldSyncWithCrm => IsPublic;

        /// <summary>
        /// Create a new instance of a SyncState, with this itemId, expected to be the EntryId of the item I wrap.
        /// </summary>
        /// <param name="itemId">the EntryId of the item I wrap</param>
        public SyncState(string itemId)
        {
            this.OutlookItemEntryId = itemId;
        }

        public void RemoveCrmLink()
        {
            CrmEntryId = null;
            if (!IsDeletedInOutlook)
            {
                RemoveSynchronisationProperties();
            }
        }

        /// <summary>
        /// Remove all synchronisation properties from this object.
        /// </summary>
        public abstract void RemoveSynchronisationProperties();

        /// <summary>
        /// Save my Outlook item.
        /// </summary>
        internal abstract void SaveItem();
    }
}
