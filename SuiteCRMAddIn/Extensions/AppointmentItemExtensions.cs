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
    using BusinessLogic;
    using Extensions;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Extension methods for Outlook AppointmentItems.
    /// </summary>
    /// <remarks>
    /// TODO: There are many methods in AppointmentSyncing which should be refactored into here.
    /// </remarks>
    public static class AppointmentItemExtension
    {
        /// <summary>
        /// True if this item represents an appointment/call.
        /// </summary>
        /// <param name="olItem">The item</param>
        /// <returns>True if this item represents an appointment/call.</returns>
        public static bool IsCall(this Outlook.AppointmentItem olItem)
        {
            return olItem.MeetingStatus == Microsoft.Office.Interop.Outlook.OlMeetingStatus.olNonMeeting;
        }

        /// <summary>
        /// Remove all the synchronisation properties from this item.
        /// </summary>
        /// <param name="olItem">The item from which the property should be removed.</param>
        public static void ClearSynchronisationProperties(this Outlook.AppointmentItem olItem)
        {
            olItem.ClearUserProperty(Synchroniser<Outlook.AppointmentItem>.CrmIdPropertyName);
            olItem.ClearUserProperty(Synchroniser<Outlook.AppointmentItem>.ModifiedDatePropertyName);
            olItem.ClearUserProperty(Synchroniser<Outlook.AppointmentItem>.TypePropertyName);
        }

        /// <summary>
        /// Removed the specified user property from this item.
        /// </summary>
        /// <param name="olItem">The item from which the property should be removed.</param>
        /// <param name="name">The name of the property to remove.</param>
        public static void ClearUserProperty(this Outlook.AppointmentItem olItem, string name)
        {
            Outlook.UserProperty olProperty = olItem.UserProperties[name];
            if (olProperty != null)
            {
                olProperty.Delete();
            }
        }

        /// <summary>
        /// Ensure that this appointment item has a recipient with this smtpAddress; if it has not,
        /// add one to it.
        /// </summary>
        /// <param name="olItem">The appointment.</param>
        /// <param name="smtpAddress">The SMTP address.</param>
        public static void EnsureRecipient(this Outlook.AppointmentItem olItem, string smtpAddress)
        {
            olItem.EnsureRecipient(smtpAddress, smtpAddress);
        }

        /// <summary>
        /// Ensure that this appointment item has a recipient with this smtpAddress; if it has not,
        /// add one with this identifier.
        /// </summary>
        /// <param name="olItem">The appointment.</param>
        /// <param name="smtpAddress">The SMTP address.</param>
        /// <param name="identifier">the identifier - which may be just the SMTP address or may be 
        /// '{SMTP address} : {phone_number}'</param>
        public static void EnsureRecipient(this Outlook.AppointmentItem olItem, string smtpAddress, string identifier)
        {
            bool found = false;

            foreach (Outlook.Recipient recipient in olItem.Recipients)
            {
                found |= recipient.GetSmtpAddress().Equals(smtpAddress);

                if (found) break;
            }

            if (!found)
            {
                olItem.Recipients.Add(identifier);
            }
        }
    }
}
