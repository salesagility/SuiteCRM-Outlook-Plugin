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
    using Extensions;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public class MeetingsSynchroniser : AppointmentSyncing
    {
        public const string CrmModule = "Meetings";

        public MeetingsSynchroniser(string name, SyncContext context) : base(name, context)
        {
        }

        public override string DefaultCrmModule
        {
            get
            {
                return CrmModule;
            }
        }

        public override SyncDirection.Direction Direction => Properties.Settings.Default.SyncMeetings;

        /// <summary>
        /// Specialisation: also set end time and location.
        /// </summary>
        /// <param name="crmItem">The CRM item.</param>
        /// <param name="olItem">The Outlook item.</param>
        protected override void SetOutlookItemDuration(EntryValue crmItem, Outlook.AppointmentItem olItem)
        {
            try
            {
                base.SetOutlookItemDuration(crmItem, olItem);
                olItem.Location = crmItem.GetValueAsString("location");
                olItem.End = olItem.Start.AddMinutes(olItem.Duration);
            }
            catch (Exception any)
            {
                Log.Error("AppointmentSyncing.SetOutlookItemDuration", any);
            }
        }

       protected override void UpdateOutlookDetails(string crmType, EntryValue crmItem, DateTime date_start, Outlook.AppointmentItem olItem)
        {
            try
            {
                olItem.Start = date_start;
                var minutesString = crmItem.GetValueAsString("duration_minutes");
                var hoursString = crmItem.GetValueAsString("duration_hours");

                int minutes = string.IsNullOrWhiteSpace(minutesString) ? 0 : int.Parse(minutesString);
                int hours = string.IsNullOrWhiteSpace(hoursString) ? 0 : int.Parse(hoursString);

                olItem.Duration = minutes + hours * 60;

                olItem.Location = crmItem.GetValueAsString("location");
                olItem.End = olItem.Start;
                if (hours > 0)
                {
                    olItem.End.AddHours(hours);
                }
                if (minutes > 0)
                {
                    olItem.End.AddMinutes(minutes);
                }
                SetRecipients(olItem, crmItem, crmItem.GetValueAsString("id"), crmType);
            }
            finally
            {
                this.SaveItem(olItem);
            }
        }


        protected override bool ShouldAddOrUpdateItemFromCrmToOutlook(Outlook.MAPIFolder folder, string crmType, EntryValue crmItem)
        {
            return crmType == "Meetings";
        }

        //internal override string AddOrUpdateItemFromOutlookToCrm(SyncState<Outlook.AppointmentItem> syncState)
        //{
        //    string result = null;
        //    if (syncState.OutlookItem.IsCall())
        //    {
        //        result = base.AddOrUpdateItemFromOutlookToCrm(syncState);
        //    }
        //    return result;
        //}
    }
}
