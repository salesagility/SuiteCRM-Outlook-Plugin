/**
 * Outlook integration for SuiteCRM.
 * @package Outlook integration for SuiteCRM
 * @copyright SalesAgility Ltd http://www.salesagility.com
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU AFFERO GENERAL PUBLIC LICENSE as published by
 * the Free Software Foundation; either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU AFFERO GENERAL PUBLIC LICENSE
 * along with this program; if not, see http://www.gnu.org/licenses
 * or write to the Free Software Foundation,Inc., 51 Franklin Street,
 * Fifth Floor, Boston, MA 02110-1301  USA
 *
 * @author SalesAgility <info@salesagility.com>
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SuiteCRMClient.RESTObjects;
using SuiteCRMClient;
using Microsoft.Office.Interop.Outlook;


namespace SuiteCRMAddIn
{
    public partial class frmAddressBook : Form
    {
        private clsSettings settings = Globals.ThisAddIn.settings;
        public frmAddressBook()
        {
            this.InitializeComponent();
            if (Globals.ThisAddIn.SuiteCRMUserSession.id == "")
            {
                frmSettings objacbbSettings = new frmSettings();
                objacbbSettings.ShowDialog();                
            }            
        }
       
        private void btnAddBCC_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in this.lstViewResults.SelectedItems)
            {
                this.lstViewBCC.Items.Add((ListViewItem) item.Clone());
            }
        }

        private void btnAddCC_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in this.lstViewResults.SelectedItems)
            {
                this.lstViewCC.Items.Add((ListViewItem) item.Clone());
            }
        }

        private void btnAddTo_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in this.lstViewResults.SelectedItems)
            {
                this.lstViewTo.Items.Add((ListViewItem) item.Clone());
            }
        }

      
        private void btnRemoveBCC_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in this.lstViewBCC.SelectedItems)
            {
                this.lstViewBCC.Items.Remove(item);
            }
        }

        private void btnRemoveCC_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in this.lstViewCC.SelectedItems)
            {
                this.lstViewCC.Items.Remove(item);
            }
        }

        private void btnRemoveTo_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in this.lstViewTo.SelectedItems)
            {
                this.lstViewTo.Items.Remove(item);
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string[] strArray = new string[2];
            string str = "OR";
            Cursor.Current = Cursors.WaitCursor;
            if (this.txtSearch.Text == string.Empty)
            {
                MessageBox.Show("Please enter something to search for", "Error");
            }
            else
            {
                this.lstViewResults.Items.Clear();
                string[] strArray2 = new string[] { "Leads", "Contacts" };
                if (this.txtSearch.Text.Contains<char>(' '))
                {
                    strArray = this.txtSearch.Text.Split(new char[] { ' ' });
                }
                else
                {
                    strArray[0] = this.txtSearch.Text;
                    strArray[1] = this.txtSearch.Text;
                }
                if ((strArray[1] != string.Empty) && (strArray[0] != strArray[1]))
                {
                    str = "AND";
                }
                foreach (string str2 in strArray2)
                {
                    string query = "(" + str2.ToLower() + ".first_name LIKE '%" + strArray[0] + "%' " + str + " " + str2.ToLower() + ".last_name LIKE '%" + strArray[1] + "%')";
                    bool flag1 = str2 == "Contacts";
                    if (this.cbMyItems.Checked)
                    {
                        string str8 = query;
                        query = str8 + "AND " + str2.ToLower() + ".assigned_user_id = '" + Globals.ThisAddIn.SuiteCRMUserSession.id + "'";
                    }
                    foreach (eEntryValue _value in clsSuiteCRMHelper.GetEntryList(str2, query, 0, "date_entered DESC", 0, false, new string[] { "first_name", "last_name", "email1" }).entry_list)
                    {
                        string str4 = string.Empty;
                        string str5 = string.Empty;
                        string valueByKey = string.Empty;
                        string str7 = string.Empty;
                        valueByKey = clsSuiteCRMHelper.GetValueByKey(_value, "first_name");
                        str7 = clsSuiteCRMHelper.GetValueByKey(_value, "last_name");
                        clsSuiteCRMHelper.GetValueByKey(_value, "id");
                        str5 = clsSuiteCRMHelper.GetValueByKey(_value, "email1");
                        str4 = valueByKey + " " + str7;
                        this.lstViewResults.Items.Add(new ListViewItem(new string[] { str4, str5, str2 }));
                    }
                }
                Cursor.Current = Cursors.Default;
            }
        }

        private void frmAddressBook_Load(object sender, EventArgs e)
        {
            this.txtSearch.KeyPress += new KeyPressEventHandler(this.frmAddressBook_KeyPress);  
        }

        private void frmAddressBook_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\r')
            {
                e.Handled = true;
                this.btnSearch_Click(null, null);
            }
        }

        private void btnFinish_Click(object sender, EventArgs e)
        {
            try
            {
                MailItem currentItem = (MailItem)Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;
                foreach (ListViewItem item2 in this.lstViewTo.Items)
                {
                    currentItem.Recipients.Add(item2.SubItems[1].Text).Type = 1;
                }
                foreach (ListViewItem item3 in this.lstViewCC.Items)
                {
                    currentItem.Recipients.Add(item3.SubItems[1].Text).Type = 2;
                }
                foreach (ListViewItem item4 in this.lstViewBCC.Items)
                {
                    currentItem.Recipients.Add(item4.SubItems[1].Text).Type = 3;
                }
                currentItem.Recipients.ResolveAll();
            }
            catch (System.Exception exception)
            {
                MessageBox.Show("Error setting address from SuiteCRM addressbook:" + exception.Message, "ERROR");
            }
            base.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            base.Close();
        }
        private void txtSearch_Enter(object sender, EventArgs e)
        {
            if (txtSearch.Focused == true)
            {
                this.AcceptButton = btnSearch;
            }
        }

        private void txtSearch_Leave(object sender, EventArgs e)
        {
            this.AcceptButton = btnFinish;
        }
    }
}
