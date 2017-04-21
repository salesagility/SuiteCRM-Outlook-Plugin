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
    using BusinessLogic;
    using Microsoft.Office.Interop.Outlook;
    using SuiteCRMClient;
    using SuiteCRMClient.Email;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows.Forms;

    public class EmailArchiveAction : AbstractDaemonAction
    {
        private readonly List<MailItem> items = new List<MailItem>();

        private readonly List<CrmEntity> entities = new List<CrmEntity>();

        private EmailArchiveType achiveType;
        private clsEmailArchive emailToArchive;
        private UserSession session;

        public EmailArchiveAction(UserSession session, clsEmailArchive objEmail, EmailArchiveType archiveType) : base(5)
        {
            this.session = session;
            this.emailToArchive = objEmail;
            this.achiveType = archiveType;
        }


        public override string Description
        {
            get
            {
                return $"Archiving {items.Count()} email items"; ;
            }
        }

        public override void Perform()
        {
            this.ReportOnEmailArchiveSuccess(
                items.Select(mailItem =>
                        Globals.ThisAddIn.EmailArchiver.ArchiveEmailWithEntityRelationships(mailItem, this.entities, this.achiveType.ToString()))
                    .ToList());
        }

        private bool ReportOnEmailArchiveSuccess(List<ArchiveResult> emailArchiveResults)
        {
            var successCount = emailArchiveResults.Count(r => r.IsSuccess);
            var failCount = emailArchiveResults.Count - successCount;
            var fullSuccess = failCount == 0;
            if (fullSuccess)
            {
                if (Globals.ThisAddIn.Settings.ShowConfirmationMessageArchive)
                {
                    MessageBox.Show(
                        $"{successCount} email(s) have been successfully archived",
                        "Archived successfully", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return true;
            }
            else
            {
                var message = successCount == 0
                    ? $"Failed to archive {failCount} email(s)"
                    : $"{successCount} emails(s) were successfully archived.";

                var first11Problems = emailArchiveResults.SelectMany(r => r.Problems).Take(11).ToList();
                if (first11Problems.Any())
                {
                    message =
                        message +
                        "\n\nThere were some failures:\n" +
                        string.Join("\n", first11Problems.Take(10)) +
                        (first11Problems.Count > 10 ? "\n[and more]" : string.Empty);
                }

                MessageBox.Show(message, "Archive failure", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }
    }
}
