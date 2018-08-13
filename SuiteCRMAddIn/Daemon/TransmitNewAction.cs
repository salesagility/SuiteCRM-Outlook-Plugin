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
    using Exceptions;
    using SuiteCRMAddIn.BusinessLogic;
    using SuiteCRMClient;
    using System.Net;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// An action to transmit to the server an item which is a new item, and
    /// does not already have a valid CRM id.
    /// </summary>
    /// <typeparam name="OutlookItemType">The type of item I transmit.</typeparam>
    public class TransmitNewAction<OutlookItemType, SyncStateType> : AbstractDaemonAction
        where OutlookItemType : class
        where SyncStateType : SyncState<OutlookItemType>
    {
        private SyncState<OutlookItemType> syncState;
        private Synchroniser<OutlookItemType, SyncStateType> synchroniser;

        public TransmitNewAction(Synchroniser<OutlookItemType, SyncStateType> synchroniser, SyncStateType state) : base(1)
        {
            /* step the state transition engine forward to queued */
            if (state.TxState == TransmissionState.NewFromOutlook)
            {
                state.SetPending();
                state.SetQueued();
            }
            this.syncState = state;
            this.synchroniser = synchroniser;
        }

        public override string Description
        {
            get
            {
                return $"{this.GetType().Name} ({this.synchroniser.DefaultCrmModule})";
            }
        }

        public override string Perform()
        {
            string result;

            /* #223: ensure that the state has a crmId that is null or empty.
             * If not null or empty then this is not a new item: do nothing and exit. */
            if (CrmId.IsInvalid(syncState.CrmEntryId))
            {
                if (syncState.TxState == TransmissionState.Queued)
                {
                    try
                    {
                        CrmId returnedCrmId = this.synchroniser.AddOrUpdateItemFromOutlookToCrm(syncState);
                        result = $"synced new item as {returnedCrmId}.\n\t{syncState.Description}";
                    }
                    catch (WebException wex)
                    {
                        if (wex.Status == WebExceptionStatus.ProtocolError)
                        {
                            using (HttpWebResponse response = wex.Response as HttpWebResponse)
                            {
                                switch (response.StatusCode)
                                {
                                    case HttpStatusCode.RequestTimeout:
                                    case HttpStatusCode.ServiceUnavailable:
                                        throw new ActionRetryableException($"Temporary error ({response.StatusCode})", wex);
                                    default:
                                        throw new ActionFailedException($"Permanent error ({response.StatusCode})", wex);
                                }
                            }
                        }
                        else
                        {
                            throw new ActionRetryableException("Temporary network error", wex);
                        }
                    }
                }
                else
                {
                    result = $"State is {syncState.TxState}; not retransmitting";
                }
            }
            else
            {
                result = $"item was already synced as {syncState.CrmEntryId}; aborted.\n{this.syncState.Description}";
            }

            return result;
        }
    }
}
