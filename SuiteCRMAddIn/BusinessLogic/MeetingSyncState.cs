
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
    using System;
    using Microsoft.Office.Interop.Outlook;
    using ProtoItems;
    using Extensions;
    using SuiteCRMClient.RESTObjects;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient;
    using Exceptions;

    /// <summary>
    /// A sync state which syncs an appointment which is a meeting.
    /// </summary>
    public class MeetingSyncState : AppointmentSyncState
    {
        public MeetingSyncState(AppointmentItem item, CrmId crmId, DateTime modifiedDate) : base(item, crmId, modifiedDate)
        {
        }

        /// <summary>
        /// The CRM type of the item I represent.
        /// </summary>
        public override string CrmType
        {
            get
            {
                return MeetingsSynchroniser.CrmModule;
            }
        }

        public override string IdentifyingFields
        {
            get
            {
                return $"subject: '{OutlookItem.Subject}'; start: '{string.Format("{0:yyyy-MM-dd HH:mm:ss}", OutlookItem.Start.ToUniversalTime())}'";
            }
        }

        /// <summary>
        /// Construct a JSON-serialisable representation of my appointment item.
        /// </summary>
        internal override ProtoItem<AppointmentItem> CreateProtoItem()
        {
            return this.VerifyItem() ?  new ProtoAppointment<MeetingSyncState>(OutlookItem) : null;
        }


        /// <summary>
        /// Specialisation: A meeting has really changed if its recipients have changed.
        /// </summary>
        /// <returns>True if I have really changed.</returns>
        internal override bool ReallyChanged()
        {
            bool result = base.ReallyChanged();

            if (result)
            {
                var cached = this.Cache as ProtoAppointment<MeetingSyncState>;
                var current = this.CreateProtoItem() as ProtoAppointment<MeetingSyncState>;

                if (cached != null && cached.Duration != current.Duration && current.Duration == 0)
                {
                    Globals.ThisAddIn.Log.Warn(
                        $"Meeting id {this.OutlookItemEntryId} (CRM id {this.CrmEntryId}) changed to zero duration");
                    throw new DurationSetToZeroException(cached.Duration);
                }
            }
            else
            {
                var cacheValues = this.Cache as ProtoAppointment<MeetingSyncState>;

                if (cacheValues == null)
                {
                    result = true;
                }
                else
                {
                    var current = this.CreateProtoItem() as ProtoAppointment<MeetingSyncState>;

                    if (cacheValues.RecipientAddresses.Count == current.RecipientAddresses.Count)
                    {
                        for (int index = 0; index < cacheValues.RecipientAddresses.Count; index ++)
                        {
                            if (cacheValues.RecipientAddresses[index] != current.RecipientAddresses[index])
                            {
                                result = true;
                                break;
                            }
                        }
                    }
                    else
                    {
                        result = true;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Get a string representing the values of the distinct fields of this crmItem, 
        /// as a final fallback for identifying an otherwise unidentifiable object.
        /// </summary>
        /// <param name="crmItem">An item received from CRM.</param>
        /// <returns>An identifying string.</returns>
        /// <see cref="SyncState{ItemType}.IdentifyingFields"/> 
        internal static string GetDistinctFields(EntryValue crmItem)
        {
            return $"subject: '{crmItem.GetValueAsString("name")}'; start: '{crmItem.GetValueAsDateTime("date_start")}'";
        }
    }
}
