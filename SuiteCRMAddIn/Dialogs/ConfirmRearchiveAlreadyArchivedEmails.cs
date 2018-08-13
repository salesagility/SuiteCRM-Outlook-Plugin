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
namespace SuiteCRMAddIn.Dialogs
{
    using Microsoft.Office.Interop.Outlook;
    using SuiteCRMAddIn.Extensions;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Linq;
    using System.Windows.Forms;
    using System.Runtime.Serialization;

    public partial class ConfirmRearchiveAlreadyArchivedEmails : Form
    {
        public ConfirmRearchiveAlreadyArchivedEmails(IEnumerable<MailItem> mails)
        {
            InitializeComponent();

            this.alreadyArchivedEmailsGrid.DataSource = new BindingSource(
                new BindingList<TableItem>(
                mails.Select(x => new TableItem(x)).ToList()),
                null);

            this.alreadyArchivedEmailsGrid.ClearSelection();
        }

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);
            this.alreadyArchivedEmailsGrid.ClearSelection();
        }

        public class TableItem
        {
            public string From { get; set; }
            public string To { get; set; }
            public DateTime Date { get; set; }
            public string Subject { get; set; }

            public TableItem(MailItem email)
            {
                List<string> recipientAddresses = new List<string>();

                foreach (Recipient recipient in email.Recipients)
                {
                    recipientAddresses.Add(recipient.GetSmtpAddress());
                }

                this.From = email.GetSenderSMTPAddress();
                this.To = string.Join(";", recipientAddresses);
                this.Date = email.SentOn;
                this.Subject = email.Subject;
            }
        }


        /// <summary>
        /// If any of these `selectedEmails` has already been archived, pop up a dialof asking whether they should be rearchived
        /// </summary>
        /// <param name="selectedEmails"></param>
        /// <returns></returns>
        public static IEnumerable<MailItem> ConfirmAlreadyArchivedEmails(IEnumerable<MailItem> selectedEmails)
        {
            IEnumerable<MailItem> archived = selectedEmails.Where(x => !string.IsNullOrEmpty(x.GetCRMEntryId()));
            IEnumerable<MailItem> result;

            if (archived.Count() > 0)
            {
                switch (new ConfirmRearchiveAlreadyArchivedEmails(archived).ShowDialog())
                {
                    case DialogResult.Yes:
                        result = selectedEmails;
                        break;
                    case DialogResult.No:
                        result = selectedEmails.Where(x => string.IsNullOrEmpty(x.GetCRMEntryId()));
                        break;
                    default:
                        throw new DialogCancelledException();
                }
            }
            else
            {
                result = selectedEmails;
            }

            return result;
        }

    }

    [Serializable]
    internal class DialogCancelledException : System.Exception
    {
        public DialogCancelledException()
        {
        }

        public DialogCancelledException(string message) : base(message)
        {
        }

        public DialogCancelledException(string message, System.Exception innerException) : base(message, innerException)
        {
        }

        protected DialogCancelledException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
