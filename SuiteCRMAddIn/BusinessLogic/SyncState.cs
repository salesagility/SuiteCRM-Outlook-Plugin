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
namespace SuiteCRMAddIn.BusinessLogic
{
    using System;
    using System.Runtime.InteropServices;
    using SuiteCRMClient.Logging;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public abstract class SyncState
    {
        private bool _wasDeleted = false;

        public abstract string CrmType { get; }

        public string CrmEntryId { get; set; }

        public DateTime OModifiedDate { get; set; }

        public bool ExistedInCrm => !string.IsNullOrEmpty(CrmEntryId);

        public bool IsPublic => OutlookItemSensitivity == Outlook.OlSensitivity.olNormal;

        public virtual bool ShouldSyncWithCrm => IsPublic;

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

        public bool IsDeletedInOutlook
        {
            get
            {
                if (_wasDeleted) return true;
                // TODO: Make this logic more robust. Perhaps check HRESULT of COMException?
                try
                {
                    // Has the side-effect of throwing an exception if the item has been deleted:
                    var entryId = OutlookItemEntryId;
                    return false;
                }
                catch (COMException com)
                {
                    Globals.ThisAddIn.Log.Debug($"Object has probably been deleted: {com.ErrorCode}, {com.Message}");
                    _wasDeleted = true;
                    return true;
                }
            }
        }

        public void RemoveCrmLink()
        {
            CrmEntryId = null;
            if (!IsDeletedInOutlook)
            {
                OutlookUserProperties["SOModifiedDate"]?.Delete();
                OutlookUserProperties["SEntryID"]?.Delete();
            }
        }
    }
}
