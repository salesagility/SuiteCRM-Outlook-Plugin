
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
    using SuiteCRMClient;
    using SuiteCRMClient.Email;
    using System.Collections.Generic;

    /// <summary>
    /// An action which archives one email.
    /// </summary>
    public class ArchiveEmailAction : AbstractDaemonAction
    {
        /// <summary>
        /// The type of archive action (i.e. is this a manual, inbound or outbound action).
        /// </summary>
        private EmailArchiveType achiveType;
        /// <summary>
        /// The single mail to be archived.
        /// </summary>
        private clsEmailArchive mail;
        /// <summary>
        /// The session in which it should be archived.
        /// </summary>
        private UserSession session;
        /// <summary>
        /// Contact ids of contacts to which this mail should be linked.
        /// </summary>
        private List<string> contacts = new List<string>();

        /// <summary>
        /// Create a new instance of an ArchiveEmailAction.
        /// </summary>
        /// <param name="session">The session in which this mail should be archived.</param>
        /// <param name="mail">The single mail to be archived.</param>
        /// <param name="archiveType">The type of archive action (i.e. is this a manual, inbound or outbound action).</param>
        /// <param name="contacts">Contact ids of contacts to which this mail should be linked.</param>
        public ArchiveEmailAction(
            UserSession session, 
            clsEmailArchive mail, 
            EmailArchiveType archiveType, 
            List<string> contacts) : base(5)
        {
            this.session = session;
            this.mail = mail;
            this.achiveType = archiveType;
            this.contacts.AddRange(contacts);
        }

        public override void Perform()
        {
            if (session.IsLoggedIn)
            {
                this.mail.SuiteCRMUserSession = session;
                this.mail.Save(this.contacts);
            }
        }
    }
}
