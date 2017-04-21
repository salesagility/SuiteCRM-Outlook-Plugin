
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
    public class ArchiveEmailAction : AbstractDaemonAction
    {
        private EmailArchiveType achiveType;
        private clsEmailArchive emailToArchive;
        private UserSession session;
        private List<string> contactIds = new List<string>();

        public ArchiveEmailAction(
            UserSession session, 
            clsEmailArchive emailToArchive, 
            EmailArchiveType archiveType, 
            List<string> contactIds) : base(5)
        {
            this.session = session;
            this.emailToArchive = emailToArchive;
            this.achiveType = archiveType;
            this.contactIds.AddRange(contactIds);
        }

        public override void Perform()
        {
            if (session.IsLoggedIn)
            {
                this.emailToArchive.SuiteCRMUserSession = session;
                this.emailToArchive.Save(this.contactIds);
            }
        }
    }
}
