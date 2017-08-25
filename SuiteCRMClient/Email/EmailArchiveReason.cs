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
namespace SuiteCRMClient.Email
{
    /// <summary>
    /// An enumeration of reasons why an email may be archived.
    /// </summary>
    public enum EmailArchiveReason
    {
        /// <summary>
        /// It's archived because the user specifically requested it (by right clicking and selecting the option).
        /// </summary>
        /// <see cref="SuiteCRMAddIn.Dialogs.ArchiveDialog"/>
        Manual = 1,

        /// <summary>
        /// It's archived because it's an inbound (received) email which matches the current criteria for archiving inbound emails.
        /// </summary>
        /// <see cref="SuiteCRMAddIn.Dialogs.EmailAccountArchiveSettingsControl.ArchiveInboundCheckbox"/> 
        /// <see cref="SuiteCRMAddIn.BusinessLogic.EmailAccountsArchiveSettings"/> 
        Inbound = 2,

        /// <summary>
        /// It's archived because it's an outbound (sent) email which matches the current criteria for archiving outbound emails.
        /// </summary>
        /// <see cref="SuiteCRMAddIn.Dialogs.EmailAccountArchiveSettingsControl.ArchiveOutboundCheckbox"/> 
        /// <see cref="SuiteCRMAddIn.BusinessLogic.EmailAccountsArchiveSettings"/> 
        Outbound = 3,

        /// <summary>
        /// It's archived because the user has specifically selected 'Send and Archive'
        /// </summary>
        /// <see cref="SuiteCRMAddIn.Menus.SuiteCRMRibbon.btnSendAndArchive_Action"/> 
        SendAndArchive = 4
    }
}
