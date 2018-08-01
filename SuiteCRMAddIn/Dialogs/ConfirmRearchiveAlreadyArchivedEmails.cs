using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SuiteCRMAddIn.Extensions;

namespace SuiteCRMAddIn.Dialogs
{
    public partial class ConfirmRearchiveAlreadyArchivedEmails : Form
    {
        private IEnumerable<MailItem> archived;

        public ConfirmRearchiveAlreadyArchivedEmails(IEnumerable<MailItem> mails) 
        {
            this.archived = mails;
            InitializeComponent();

            this.alreadyArchivedEmailsGrid.DataSource = mails.Select(x => new TableItem(x));
        }

        class TableItem
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

                this.From = email.SenderEmailAddress;
                this.To = string.Join(";", recipientAddresses);
                this.Date = email.SentOn;
                this.Subject = email.Subject;
            }
        }
    }
}
