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
    using SuiteCRMClient;
    using System;
    using System.Globalization;
    using System.Runtime.InteropServices;
    using System.Xml.Serialization;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Extension methods for Outlook ContactItems.
    /// </summary>
    /// <remarks>
    /// TODO: There are many methods in ContactSyncing which should be refactored into here.
    /// </remarks>
    public static class ContactItemExtensions
    {
        /// <summary>
        /// Name of the override property.
        /// </summary>
        private const string OverridePropertyName = "UserOverride";

        /// <summary>
        /// The duration of the override window in minutes.
        /// </summary>
        private const int OverrideWindowMinutes = 10;

        /// <summary>
        /// Remove all the synchronisation properties from this item.
        /// </summary>
        /// <param name="olItem">The item from which the property should be removed.</param>
        public static void ClearSynchronisationProperties(this Outlook.ContactItem olItem)
        {
            olItem.ClearUserProperty(SyncStateManager.CrmIdPropertyName);
            olItem.ClearUserProperty(SyncStateManager.ModifiedDatePropertyName);
            olItem.ClearUserProperty(SyncStateManager.TypePropertyName);
        }

        /// <summary>
        /// Removed the specified user property from this item.
        /// </summary>
        /// <param name="olItem">The item from which the property should be removed.</param>
        /// <param name="name">The name of the property to remove.</param>
        public static void ClearUserProperty(this Outlook.ContactItem olItem, string name)
        {
            olItem.UserProperties[name]?.Delete();
        }


        /// <summary>
        /// Get the CRM id for this item, if known, else the empty string.
        /// </summary>
        /// <param name="olItem">The Outlook item under consideration.</param>
        /// <returns>the CRM id for this item, if known, else the empty string.</returns>
        public static CrmId GetCrmId(this Outlook.ContactItem olItem)
        {
            Outlook.UserProperty property = olItem.UserProperties[SyncStateManager.CrmIdPropertyName];

            if (property == null)
            {
                /* #6661: fail over to legacy property name if current property 
                 * name not found */
                property = olItem.UserProperties[SyncStateManager.LegacyCrmIdPropertyName];
            }

            CrmId result = property != null ? CrmId.Get(property.Value) : CrmId.Empty;

            return result;
        }

        /// <summary>
        /// True if the override window is open for this item.
        /// </summary>
        /// <remarks>In order to allow manual sync, we need to be able to override the disablement of syncing -
        /// but only briefly.</remarks>
        /// <param name="olItem">The item which we wish to sync.</param>
        /// <returns>True if the manual sync window is open for this item.</returns>
        public static bool IsManualOverride(this Outlook.ContactItem olItem)
        {
            bool result = false;
            if (olItem.UserProperties[OverridePropertyName] != null)
            {
                DateTime value = olItem.UserProperties[OverridePropertyName].Value;

                if ((DateTime.UtcNow - value).Minutes < OverrideWindowMinutes)
                {
                    result = true;
                }
                else
                {
                    /* no point holding on to a timed-out manual override property */
                    olItem.ClearManualOverride();
                }                
            }

            return result;
        }

        /// <summary>
        /// Set this item as manually syncable, briefly. As a side effect of making the change triggers sync.
        /// </summary>
        /// <remarks>In order to allow manual sync, we need to be able to override the disablement of syncing -
        /// but only briefly.</remarks>
        /// <param name="olItem">The item which may be synced despite syncing being disabled</param>
        public static void SetManualOverride(this Outlook.ContactItem olItem)
        {
            var p = olItem.UserProperties.Add(OverridePropertyName, Outlook.OlUserPropertyType.olDateTime);
            p.Value = DateTime.UtcNow;
            olItem.Save();
        }

        /// <summary>
        /// Clear the manually syncability of this item; does not break is manual sync was not set.
        /// </summary>
        /// <remarks>In order to allow manual sync, we need to be able to override the disablement of syncing -
        /// but only briefly.</remarks>
        /// <param name="olItem">The item which may be synced despite syncing being disabled</param>
        public static void ClearManualOverride(this Outlook.ContactItem olItem)
        {
            olItem.UserProperties[OverridePropertyName]?.Delete();
        }

        public static void ClearCrmId(this Outlook.ContactItem olItem)
        {
            var state = SyncStateManager.Instance.GetExistingSyncState(olItem);

            olItem.ClearUserProperty(SyncStateManager.CrmIdPropertyName);

            if (state != null)
            {
                state.CrmEntryId = null;
            }

            olItem.Save();
        }

        public static void ChangeCrmId(this Outlook.ContactItem olItem, string text)
        {
            var crmId = new CrmId(text);
            var state = SyncStateManager.Instance.GetExistingSyncState(olItem);
            var userProperty = olItem.UserProperties.Find(SyncStateManager.CrmIdPropertyName) ??
                               olItem.UserProperties.Add(SyncStateManager.CrmIdPropertyName,
                                   Outlook.OlUserPropertyType.olText);
            userProperty.Value = crmId.ToString();

            if (state != null)
            {
                state.CrmEntryId = crmId;
            }

            olItem.Save();
        }


        /// <summary>
        /// Am I actually a valid Outlook item at all?
        /// </summary>
        /// <param name="item">The item</param>
        /// <returns>True if the item is a valid COM object representing an AppointmentItem.</returns>
        public static bool IsValid(this Outlook.ContactItem item)
        {
            bool result;
            try
            {
                result = !string.IsNullOrEmpty(item.EntryID);
            }
            catch (COMException)
            {
                result = false;
            }

            return result;
        }
    }
}
