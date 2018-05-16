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
namespace SuiteCRMAddIn.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public static class MeetingItemExtensions
    {
        /// <summary>
        /// Extract the vCal uid from this MeetingItem, if available.
        /// </summary>
        /// <param name="item">A meeting item, which may relate to a meeting in CRM.</param>
        /// <returns>The vCal id if present, else the empty string.</returns>
        public static string GetVCalId(this Outlook.MeetingItem item)
        {
            Outlook.AppointmentItem appt = item.GetAssociatedAppointment(false);
            string result = string.Empty;

            //This parses the Global Appointment ID to a byte array. We need to retrieve the "UID" from it (if available).
            byte[] bytes = (byte[])appt.PropertyAccessor.StringToBinary(appt.GlobalAppointmentID);

            //According to https://msdn.microsoft.com/en-us/library/ee157690(v=exchg.80).aspx we don't need first 40 bytes            
            if (bytes.Length >= 40)
            {
                byte[] bytesThatContainData = new byte[bytes.Length - 40];
                Array.Copy(bytes, 40, bytesThatContainData, 0, bytesThatContainData.Length);

                //In some cases, there won't be a UID.
                var decoded = Encoding.UTF8.GetString(bytesThatContainData, 0, bytesThatContainData.Length);

                if (decoded.StartsWith("vCal-Uid"))
                {
                    //remove vCal-Uid from start string and special symbols
                    result = decoded.Replace("vCal-Uid", string.Empty).Replace("\u0001", string.Empty).Replace("\0", string.Empty);
                }
                else
                {
                    // Bad format!!!
                }
            }

            return result;
        }
    }
}
