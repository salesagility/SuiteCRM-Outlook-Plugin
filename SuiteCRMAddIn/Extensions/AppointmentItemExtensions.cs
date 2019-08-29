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
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
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
        /// <exception cref="COMException">Shouldn't happen, but it does, and 
        /// needs to be trapped for.</exception> 
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
            /* it doesn't actually matter whether we use CallSyncState or MeetingSyncState here,
             * since the constants are stored on Synchroniser */
            olItem.ClearUserProperty(SyncStateManager.CrmIdPropertyName);
            olItem.ClearUserProperty(SyncStateManager.ModifiedDatePropertyName);
            olItem.ClearUserProperty(SyncStateManager.TypePropertyName);
        }


        /// <summary>
        /// Get the CRM id for this item, if known, else the empty string.
        /// </summary>
        /// <param name="olItem">The Outlook item under consideration.</param>
        /// <returns>the CRM id for this item, if known, else null.</returns>
        public static CrmId GetCrmId(this Outlook.AppointmentItem olItem)
        {
            string result;

            if (olItem.IsValid())
            {
                try
                {
                    Outlook.UserProperty property = olItem.UserProperties[SyncStateManager.CrmIdPropertyName];

                    if (property == null)
                    {
                        /* #6661: fail over to legacy property name if current property 
                         * name not found */
                        property = olItem.UserProperties[SyncStateManager.LegacyCrmIdPropertyName];
                    }
                    
                    if (property != null && !string.IsNullOrEmpty(property.Value))
                    {
                        result = property.Value;
                    }
                    else
                    {
                        result = olItem.GetVCalId();
                    }
                } 
                catch (COMException)
                {
                    /* this is bad! It shouldn't be possible to get here, but 
                     * it is. */
                    try
                    {
                        result = olItem.GetVCalId();
                    }
                    catch (COMException)
                    {
                        result = null;
                    }
                }
            }
            else
            {
                result = null;
            }

            return CrmId.Get(result);
        }

        /// <summary>
        /// Set the CRM id for this item to this value.
        /// </summary>
        /// <param name="olItem">The Outlook item under consideration.</param>
        /// <param name="crmId">The value to set.</param>
        public static void SetCrmId(this Outlook.AppointmentItem olItem, CrmId crmId)
        {
            Outlook.UserProperty property = olItem.UserProperties[SyncStateManager.CrmIdPropertyName];

            if (property == null)
            {
                property = olItem.UserProperties.Add(SyncStateManager.CrmIdPropertyName, Outlook.OlUserPropertyType.olText);
                SyncStateManager.Instance.SetByCrmId(crmId, SyncStateManager.Instance.GetOrCreateSyncState(olItem));
            }
            if (CrmId.IsInvalid(crmId))
            {
                property.Delete();
            }
            else
            {
                property.Value = crmId.ToString();
            }
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

        /// <summary>
        /// Extract the vCal uid from this AppointmentItem, if available.
        /// </summary>
        /// <param name="item">An appointment item, which may relate to a meeting in CRM.</param>
        /// <returns>The vCal id if present, else the empty string.</returns>
        public static string GetVCalId(this Outlook.AppointmentItem item)
        {
            string result = string.Empty;

            //This parses the Global Appointment ID to a byte array. We need to retrieve the "UID" from it (if available).
            byte[] bytes = (byte[])item.PropertyAccessor.StringToBinary(item.GlobalAppointmentID);

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
#if DEBUG
                else
                {
                    // Bad format!
                    Globals.ThisAddIn.Log.Debug($"Failed to find vCal-Uid in GlobalAppointmentId '{Encoding.UTF8.GetString(bytes)}' in appointment '{item.Subject}'");
                }
#endif
            }
            else
            {
                Globals.ThisAddIn.Log.Debug($"Failed to find vCal-Uid in short GlobalAppointmentId '{Encoding.UTF8.GetString(bytes)}' in appointment '{item.Subject}'");
            }

            return result;
        }

        /// <summary>
        /// Am I actually a valid Outlook item at all?
        /// </summary>
        /// <param name="item">The item</param>
        /// <returns>True if the item is a valid COM object representing an AppointmentItem.</returns>
        public static bool IsValid(this Outlook.AppointmentItem item)
        {
            bool result;
            try
            {
                result = !string.IsNullOrEmpty(item.EntryID);
            }
            catch (Exception e) when (e is COMException || e is NullReferenceException)
            {
                result = false;
            }

            return result;
        }
    }
}
