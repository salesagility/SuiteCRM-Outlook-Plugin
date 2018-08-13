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

        public bool IsDeletedInOutlook
        {
            get
            {
                bool result;
                if (_wasDeleted) return true;
                // TODO: Make this logic more robust. Perhaps check HRESULT of COMException?
                try
                {
                    // Has the side-effect of throwing an exception if the item has been deleted:
                    var entryId = OutlookItemEntryId;
                    result = false;
                }
                catch (COMException com)
                {
                    Globals.ThisAddIn.Log.Debug($"Object has probably been deleted: {com.ErrorCode}, {com.Message}; HResult {com.HResult}");
                    _wasDeleted = true;
                    result = true;
                }

                return result;
            }
        }

        public bool IsPublic => OutlookItemSensitivity == Outlook.OlSensitivity.olNormal;

        public DateTime OModifiedDate { get; set; }

        /// <summary>
        /// Precisely 'this.OutlookItem.EntryId'.
        /// </summary>
        /// <remarks>Outlook item classes do not inherit from a common base class, so generic client code cannot refer to 'OutlookItem.EntryId'.</remarks>
        public abstract string OutlookItemEntryId { get; }

        /// <summary>
        /// Precisely 'this.OutlookItem.Sensitivity'.
        /// </summary>
        /// <remarks>Outlook item classes do not inherit from a common base class, so generic client code cannot refer to 'OutlookItem.Sensitivity'.</remarks>
        public abstract Outlook.OlSensitivity OutlookItemSensitivity { get; }

        public abstract Outlook.UserProperties OutlookUserProperties { get; }

        public virtual bool ShouldSyncWithCrm => IsPublic;

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
