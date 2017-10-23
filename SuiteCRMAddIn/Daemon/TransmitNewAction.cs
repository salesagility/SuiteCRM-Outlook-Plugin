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
    using SuiteCRMAddIn.BusinessLogic;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// An action to transmit to the server an item which is a new item, and
    /// does not already have a valid CRM id.
    /// </summary>
    /// <typeparam name="OutlookItemType">The type of item I transmit.</typeparam>
    public class TransmitNewAction<OutlookItemType> : AbstractDaemonAction
        where OutlookItemType : class
    {
        private string crmType;
        private SyncState<OutlookItemType> syncState;
        private Synchroniser<OutlookItemType> synchroniser;

        public TransmitNewAction(Synchroniser<OutlookItemType> synchroniser, SyncState<OutlookItemType> state, string crmType) : base(1)
        {
            this.syncState = state;
            this.crmType = crmType;
            this.synchroniser = synchroniser;
        }

        public override string Description
        {
            get
            {
                return $"{this.GetType().Name} ({this.crmType})";
            }
        }

        public override string Perform()
        {
            /* #223: ensure that the state has a crmId is null or empty.
             * If not null or empty then this is not a new item: do nothing and exit. */
            if (string.IsNullOrEmpty(syncState.CrmEntryId))
            {
                return $"synced new item as {this.synchroniser.AddOrUpdateItemFromOutlookToCrm(syncState, this.crmType)}.\n{this.syncState.Description}";
            }
            else
            {
                return $"item was already synced as {syncState.CrmEntryId}; aborted.\n{this.syncState.Description}";
            }
        }
    }
}
