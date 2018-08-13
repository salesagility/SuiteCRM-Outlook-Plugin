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
    using SuiteCRMClient.RESTObjects;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public class CallsSynchroniser : AppointmentsSynchroniser<CallSyncState>
    {
        public const string CrmModule = "Calls";

        public CallsSynchroniser(string name, SyncContext context) : base(name, context)
        {
        }

        public override string DefaultCrmModule
        {
            get
            {
                return CrmModule;
            }
        }

        public override SyncDirection.Direction Direction => Properties.Settings.Default.SyncCalls;

        protected override void InstallEventHandlers()
        {
            /* arbitrarily, one AppointmentSyncing subclass should NOT handle events. */
        }

        protected override void RemoveEventHandlers()
        {
            /* arbitrarily, one AppointmentSyncing subclass should NOT handle events. */
        }

        protected override void SetMeetingStatus(Outlook.AppointmentItem olItem, EntryValue crmItem)
        {
            olItem.MeetingStatus = Microsoft.Office.Interop.Outlook.OlMeetingStatus.olNonMeeting;
        }

        protected override bool ShouldAddOrUpdateItemFromCrmToOutlook(Outlook.MAPIFolder folder, string crmType, EntryValue crmItem)
        {
            return crmType == "Calls";
        }
    }
}
