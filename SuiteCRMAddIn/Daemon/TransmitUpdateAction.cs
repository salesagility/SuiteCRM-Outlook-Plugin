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
namespace SuiteCRMAddIn.Daemon
{
    using BusinessLogic;
    using Exceptions;
    using ProtoItems;
    using SuiteCRMClient.Logging;
    using System.Net;
    using System.Runtime.InteropServices;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// An action to transmit to the server an item which is not a new item, but
    /// already has a SyncState.
    /// </summary>
    /// <typeparam name="OutlookItemType">The type of item I transmit.</typeparam>
    public class TransmitUpdateAction<OutlookItemType, SyncStateType> : AbstractDaemonAction
        where OutlookItemType : class
        where SyncStateType : SyncState<OutlookItemType>
    {
        /// <summary>
        /// The synchroniser I need to call to perform this action.
        /// </summary>
        private Synchroniser<OutlookItemType, SyncStateType> synchroniser;
        /// <summary>
        /// The sync state on which this action should be performed.
        /// </summary>
        private SyncState<OutlookItemType> state;


        /// <summary>
        /// Create a new instance of the TrensmitUpdateItem class, wrapping this state.
        /// </summary>
        /// <param name="synchroniser">The synchroniser I will call to perform this action.</param>
        /// <param name="state">The sync state on which this action should be performed.</param>
        public TransmitUpdateAction(Synchroniser<OutlookItemType, SyncStateType> synchroniser, SyncStateType state) : base(1)
        {
            state.SetQueued();
            this.synchroniser = synchroniser;
            this.state = state;

            MeetingSyncState meeting = state as MeetingSyncState;

            if (meeting != null)
            {
                ILogger log = Globals.ThisAddIn.Log;
                try
                {
                    switch (meeting.Cache.Status) {
                        case Outlook.OlMeetingStatus.olMeetingCanceled:
                            log.Info($"TransmitUpdateAction: registered meeting {state.Description} cancelled");
                            break;
                        case Microsoft.Office.Interop.Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled:
                            log.Info($"TransmitUpdateAction: registered meeting {state.Description} received and cancelled");
                            break;
                    }
                }
                catch (COMException comx)
                {
                    ErrorHandler.Handle($"Possibly-deleted item while trying to transmit update? HResult = {comx.HResult}", comx);
                }
            }
        }


        public override string Description
        {
            get
            {
                try
                {
                    return $"{this.GetType().Name} ({state.CrmType} {state.Description})";
                }
                catch (COMException comx)
                {
                    ErrorHandler.Handle($"Possibly-deleted item while trying to transmit update? HResult = {comx.HResult}", comx);
                    return $"{this.GetType().Name} ({state.CrmType} - possibly cancelled meeting?";
                }
            }
        }


        public override string Perform()
        {
            try
            {
                try
                {
                    var id = state.CrmEntryId;
                    if (!synchroniser.AddOrUpdateItemFromOutlookToCrm(state).IsValid())
                    {
                        throw new ActionRetryableException($"Unexpected response from CRM while attempting to sync item {id}");
                    }
                }
                catch (COMException comx)
                {
                    ErrorHandler.Handle($"Possibly-deleted item while trying to transmit update? HResult = {comx.HResult}", comx);
                    synchroniser.HandleItemMissingFromOutlook(state);
                }

                return "Synced.";
            }
            catch (WebException wex)
            {
                throw new ActionRetryableException("Temporary network error", wex);
            }
        }
    }
}
