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
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading.Tasks;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Extension methods for Outlook TaskItems.
    /// </summary>
    /// <remarks>
    /// TODO: There are many methods in TaskSyncing which should be refactored into here.
    /// </remarks>
    public static class TaskItemExtensions
    {
        /// <summary>
        /// Remove all the synchronisation properties from this item.
        /// </summary>
        /// <param name="olItem">The item from which the property should be removed.</param>
        public static void ClearSynchronisationProperties(this Outlook.TaskItem olItem)
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
        public static void ClearUserProperty(this Outlook.TaskItem olItem, string name)
        {
            Outlook.UserProperty olProperty = olItem.UserProperties[name];
            if (olProperty != null)
            {
                olProperty.Delete();
            }
        }


        /// <summary>
        /// Get the CRM id for this item, if known, else the empty string.
        /// </summary>
        /// <param name="olItem">The Outlook item under consideration.</param>
        /// <returns>the CRM id for this item, if known, else the empty string.</returns>
        public static CrmId GetCrmId(this Outlook.TaskItem olItem)
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
        /// Am I actually a valid Outlook item at all?
        /// </summary>
        /// <param name="item">The item</param>
        /// <returns>True if the item is a valid COM object representing an AppointmentItem.</returns>
        public static bool IsValid(this Outlook.TaskItem item)
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
