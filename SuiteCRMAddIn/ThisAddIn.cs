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
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using SuiteCRMClient;
using System.Runtime.InteropServices;
using SuiteCRMAddIn.Properties;
using System.Security.Cryptography;
using System.Globalization;
using SuiteCRMClient.RESTObjects;
using SuiteCRMClient;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SuiteCRMAddIn
{
    public partial class ThisAddIn
    {
        public SuiteCRMClient.clsUsersession SuiteCRMUserSession;
        public clsSettings settings;
        private Outlook.Explorer objExplorer;
        public Office.CommandBarPopup objSuiteCRMMenuBar2007;
        public Office.CommandBarButton btnArvive;
        public Office.CommandBarButton btnSettings;
        List<Outlook.Folder> lstOutlookFolders;
        public int CurrentVersion;
        List<cAppItem> lCalItems;
        private string sDelCalId = "";
        private string sDelCalModule = "";
        private bool SyncInProgress = false;
        private bool IsCalendarView = false;
        private string PrevCalSID = "";

        List<cTaskItem> lTaskItems;
        private string sDelTaskId = "";
        private bool IsTaskView = false;
        private string PrevTaskID = "";

        List<cContactItem> lContactItems;
        private string sDelContactId = "";
        private bool IsContactView = false;
        private string PrevContactID = "";
        public Office.IRibbonUI RibbonUI { get; set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            CurrentVersion = Convert.ToInt32(Globals.ThisAddIn.Application.Version.Split('.')[0]);
            this.objExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
            SuiteCRMClient.clsSuiteCRMHelper.InstallationPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SuiteCRMOutlookAddIn";
            this.settings = new clsSettings();
            this.objExplorer.FolderSwitch -= objExplorer_FolderSwitch;
            this.objExplorer.FolderSwitch += objExplorer_FolderSwitch;
            if (this.settings.AutoArchive)
            {
                this.objExplorer.Application.NewMailEx += new Outlook.ApplicationEvents_11_NewMailExEventHandler(this.Application_NewMail);
                this.objExplorer.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(this.Application_ItemSend);
            }
            if (CurrentVersion < 14)
            {
                this.Application.ItemContextMenuDisplay += new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(this.Application_ItemContextMenuDisplay);
                var menuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;
                objSuiteCRMMenuBar2007 = (Office.CommandBarPopup)menuBar.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, true);
                if (objSuiteCRMMenuBar2007 != null)
                {
                    objSuiteCRMMenuBar2007.Caption = "SuiteCRM";
                    this.btnArvive = (Office.CommandBarButton)this.objSuiteCRMMenuBar2007.Controls.Add(Office.MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                    this.btnArvive.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    this.btnArvive.Caption = "Archive";
                    this.btnArvive.Picture = RibbonImageHelper.Convert(Resources.SuiteCRM1);
                    this.btnArvive.Click += new Office._CommandBarButtonEvents_ClickEventHandler(this.cbtnArchive_Click);
                    this.btnArvive.Visible = true;
                    this.btnArvive.BeginGroup = true;
                    this.btnArvive.TooltipText = "Archive selected emails to SuiteCRM";
                    this.btnArvive.Enabled = true;
                    this.btnSettings = (Office.CommandBarButton)this.objSuiteCRMMenuBar2007.Controls.Add(Office.MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                    this.btnSettings.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    this.btnSettings.Caption = "Settings";
                    this.btnSettings.Click += new Office._CommandBarButtonEvents_ClickEventHandler(this.cbtnSettings_Click);
                    this.btnSettings.Visible = true;
                    this.btnSettings.BeginGroup = true;
                    this.btnSettings.TooltipText = "SuiteCRM Settings";
                    this.btnSettings.Enabled = true;
                    this.btnSettings.Picture = RibbonImageHelper.Convert(Resources.Settings);            

                    objSuiteCRMMenuBar2007.Visible = true;
                }
            }
            else
            {
                //For Outlook version 2010 and greater
                //var app = this.Application;
                //app.FolderContextMenuDisplay += new Outlook.ApplicationEvents_11_FolderContextMenuDisplayEventHander(this.app_FolderContextMenuDisplay);
            }
            SuiteCRMAuthenticate();
            Sync();
        }

        void objExplorer_FolderSwitch()
        {
            if (this.objExplorer.CurrentFolder.DefaultItemType == GetDefaultFolder("appointments").DefaultItemType)
                IsCalendarView = true;
            else
                IsCalendarView = false;
            if (this.objExplorer.CurrentFolder.DefaultItemType == GetDefaultFolder("tasks").DefaultItemType)
                IsTaskView = true;
            else
                IsTaskView = false;
            if (this.objExplorer.CurrentFolder.DefaultItemType == GetDefaultFolder("contacts").DefaultItemType)
                IsContactView = true;
            else
                IsContactView = false;
        }

        public async void Sync()
        //public void Sync()
        {
            while (1 == 1)
            {
                if (SuiteCRMUserSession != null && SuiteCRMUserSession.id != "") 
                {
                    if (settings.SyncCalendar)
                    {
                        //StartCalendarSync();
                        Thread oThread = new Thread(() => StartCalendarSync());
                        oThread.Start();
                        //StartTaskSync();
                        Thread oThread1 = new Thread(() => StartTaskSync());
                        oThread1.Start();                        
                    }
                    if (settings.SyncContacts)
                    {
                        //StartContactSync();
                        Thread oThread = new Thread(() => StartContactSync());
                        oThread.Start(); 
                    }
                }
                await Task.Delay(300000); //5 mins delay
            }
        }
        private void StartContactSync()
        {
            Outlook.NameSpace oNS = this.Application.GetNamespace("mapi");
            Outlook.MAPIFolder contactsFolder = GetDefaultFolder("contacts");
            Outlook.Items items = contactsFolder.Items;

            items.ItemAdd -= CItems_ItemAdd;
            items.ItemChange -= CItems_ItemChange;
            items.ItemRemove -= CItems_ItemRemove;
            items.ItemAdd += CItems_ItemAdd;
            items.ItemChange += CItems_ItemChange;
            items.ItemRemove += CItems_ItemRemove;

            Outlook.MAPIFolderEvents_12_Event oCalendarEvents;
            oCalendarEvents = contactsFolder as Outlook.MAPIFolderEvents_12_Event;
            if (oCalendarEvents != null)
            {
                oCalendarEvents.BeforeItemMove -= CoCalendarEvents_BeforeItemMove;
                oCalendarEvents.BeforeItemMove += CoCalendarEvents_BeforeItemMove;
            }

            SyncInProgress = true;
            GetOutlookCItems(contactsFolder);
            SyncContacts(contactsFolder);
            SyncInProgress = false;

        }
        private void SyncContacts(Outlook.MAPIFolder contactFolder)
        {
            eGetEntryListResult _result2 = clsSuiteCRMHelper.GetEntryList("Contacts", "",
                            0, "date_entered DESC", 0, false, clsSuiteCRMHelper.GetSugarFields("Contacts"));
            if (_result2 != null)
            {
                foreach (var oResult in _result2.entry_list)
                {
                    try
                    {
                        dynamic dResult = JsonConvert.DeserializeObject(oResult.name_value_object.ToString());
                        var oItem = lContactItems.Where(a => a.SEntryID == dResult.id.value.ToString()).FirstOrDefault();
                        if (oItem == default(cContactItem))
                        {
                            Outlook.ContactItem cItem = contactFolder.Items.Add(Outlook.OlItemType.olContactItem);
                            cItem.FirstName = dResult.first_name.value.ToString();
                            cItem.LastName = dResult.last_name.value.ToString();
                            cItem.Email1Address = dResult.email1.value.ToString();
                            cItem.BusinessTelephoneNumber = dResult.phone_work.value.ToString();
                            cItem.HomeTelephoneNumber = dResult.phone_home.value.ToString();
                            cItem.MobileTelephoneNumber = dResult.phone_mobile.value.ToString();
                            cItem.JobTitle = dResult.title.value.ToString();
                            cItem.Department = dResult.department.value.ToString();
                            cItem.BusinessAddressCity = dResult.primary_address_city.value.ToString();
                            cItem.BusinessAddressCountry = dResult.primary_address_country.value.ToString();
                            cItem.BusinessAddressPostalCode = dResult.primary_address_postalcode.value.ToString();
                            cItem.BusinessAddressState = dResult.primary_address_state.value.ToString();
                            cItem.BusinessAddressStreet = dResult.primary_address_street.value.ToString();
                            cItem.Body = dResult.description.value.ToString();
                            if (dResult.account_name != null)
                            {
                                cItem.Account = dResult.account_name.value.ToString();
                                cItem.CompanyName = dResult.account_name.value.ToString();
                            }
                            cItem.BusinessFaxNumber = dResult.phone_fax.value.ToString();
                            cItem.Title = dResult.salutation.value.ToString();

                            Outlook.UserProperty oProp = cItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                            oProp.Value = dResult.date_modified.value.ToString();
                            Outlook.UserProperty oProp2 = cItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                            oProp2.Value = dResult.id.value.ToString();
                            cItem.Save();
                            lContactItems.Add(new cContactItem
                            {
                                oItem = cItem,
                                OModifiedDate = dResult.date_modified.value.ToString(),
                                SEntryID = dResult.id.value.ToString(),
                                Touched = true
                            });
                        }
                        else
                        {
                            oItem.Touched = true;
                            Outlook.ContactItem cItem = oItem.oItem;
                            Outlook.UserProperty oProp = cItem.UserProperties["SOModifiedDate"];

                            if (oProp.Value != dResult.date_modified.value.ToString())
                            {
                                cItem.FirstName = dResult.first_name.value.ToString();
                                cItem.LastName = dResult.last_name.value.ToString();
                                cItem.Email1Address = dResult.email1.value.ToString();
                                cItem.BusinessTelephoneNumber = dResult.phone_work.value.ToString();
                                cItem.HomeTelephoneNumber = dResult.phone_home.value.ToString();
                                cItem.MobileTelephoneNumber = dResult.phone_mobile.value.ToString();
                                cItem.JobTitle = dResult.title.value.ToString();
                                cItem.Department = dResult.department.value.ToString();
                                cItem.BusinessAddressCity = dResult.primary_address_city.value.ToString();
                                cItem.BusinessAddressCountry = dResult.primary_address_country.value.ToString();
                                cItem.BusinessAddressPostalCode = dResult.primary_address_postalcode.value.ToString();
                                cItem.BusinessAddressState = dResult.primary_address_state.value.ToString();
                                cItem.BusinessAddressStreet = dResult.primary_address_street.value.ToString();
                                cItem.Body = dResult.description.value.ToString();
                                cItem.Account = dResult.account_name.value.ToString();
                                cItem.CompanyName = dResult.account_name.value.ToString();
                                cItem.BusinessFaxNumber = dResult.phone_fax.value.ToString();
                                cItem.Title = dResult.salutation.value.ToString();
                                if (oProp == null)
                                    oProp = cItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                                oProp.Value = dResult.date_modified.value.ToString();
                                Outlook.UserProperty oProp2 = cItem.UserProperties["SEntryID"];
                                if (oProp2 == null)
                                    oProp2 = cItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                                oProp2.Value = dResult.id.value.ToString();
                                cItem.Save();
                            }
                        }
                    }
                    catch
                    { }
                }
            }

            try
            {
                var lItemToBeDeletedO = lContactItems.Where(a => !a.Touched && !string.IsNullOrWhiteSpace(a.OModifiedDate)).ToList();
                foreach (var oItem in lItemToBeDeletedO)
                {
                    oItem.oItem.Delete();
                }
                lContactItems.RemoveAll(a => !a.Touched && !string.IsNullOrWhiteSpace(a.OModifiedDate));

                var lItemToBeAddedToS = lContactItems.Where(a => !a.Touched && string.IsNullOrWhiteSpace(a.OModifiedDate)).ToList();
                foreach (var oItem in lItemToBeAddedToS)
                {
                    AddContactToS(oItem.oItem);
                }
            }
            catch
            { }
        }
        private void GetOutlookCItems(Outlook.MAPIFolder taskFolder)
        {
            if (lContactItems == null)
            {
                lContactItems = new List<cContactItem>();
                Outlook.Items items = taskFolder.Items.Restrict("[MessageClass] = 'IPM.Contact'");
                foreach (Outlook.ContactItem oItem in items)
                {
                    if (oItem.Sensitivity != Outlook.OlSensitivity.olPrivate)
                    {
                        Outlook.UserProperty oProp = oItem.UserProperties["SOModifiedDate"];
                        if (oProp != null)
                        {
                            Outlook.UserProperty oProp2 = oItem.UserProperties["SEntryID"];
                            lContactItems.Add(new cContactItem
                            {
                                oItem = oItem,
                                OModifiedDate = oProp.Value.ToString(),
                                SEntryID = oProp2.Value.ToString()
                            });
                        }
                        else
                        {
                            lContactItems.Add(new cContactItem
                            {
                                oItem = oItem
                            });
                        }
                    }
                }
            }
        }

        void CItems_ItemChange(object Item)
        {
            if (!SyncInProgress && IsContactView)
            {
                bool SyncOldStatus = SyncInProgress;
                SyncInProgress = true;
                var oItem = Item as Outlook.ContactItem;
                if (PrevContactID != oItem.EntryID)
                {
                    Outlook.UserProperty oProp1 = oItem.UserProperties["SEntryID"];
                    if (oProp1 != null)
                    {
                        AddContactToS(oItem, oProp1.Value.ToString());
                    }
                }                
                SyncInProgress = SyncOldStatus;
            }
        }

        void CItems_ItemAdd(object Item)
        {
            if (!SyncInProgress && IsContactView)
            {
                AddContactToS(Item as Outlook.ContactItem);
            }
        }
        private void AddContactToS(Outlook.ContactItem oItem, string sID = "")
        {
            if (oItem != null)
            {
                try
                {
                    PrevContactID = oItem.EntryID;
                    string _result = "";
                    eNameValue[] data = new eNameValue[18];

                    data[0] = clsSuiteCRMHelper.SetNameValuePair("email1", oItem.Email1Address);
                    data[1] = clsSuiteCRMHelper.SetNameValuePair("title", oItem.JobTitle);
                    data[2] = clsSuiteCRMHelper.SetNameValuePair("phone_work", oItem.BusinessTelephoneNumber);
                    data[3] = clsSuiteCRMHelper.SetNameValuePair("phone_home", oItem.HomeTelephoneNumber);
                    data[4] = clsSuiteCRMHelper.SetNameValuePair("phone_mobile", oItem.MobileTelephoneNumber);
                    data[5] = clsSuiteCRMHelper.SetNameValuePair("phone_fax", oItem.BusinessFaxNumber);
                    data[6] = clsSuiteCRMHelper.SetNameValuePair("department", oItem.Department);
                    data[7] = clsSuiteCRMHelper.SetNameValuePair("primary_address_city", oItem.BusinessAddressCity);
                    data[8] = clsSuiteCRMHelper.SetNameValuePair("primary_address_state", oItem.BusinessAddressState);
                    data[9] = clsSuiteCRMHelper.SetNameValuePair("primary_address_postalcode", oItem.BusinessAddressPostalCode);
                    data[10] = clsSuiteCRMHelper.SetNameValuePair("primary_address_country", oItem.BusinessAddressCountry);
                    data[11] = clsSuiteCRMHelper.SetNameValuePair("primary_address_street", oItem.BusinessAddressStreet);
                    data[12] = clsSuiteCRMHelper.SetNameValuePair("description", oItem.Body);
                    data[13] = clsSuiteCRMHelper.SetNameValuePair("last_name", oItem.LastName);
                    data[14] = clsSuiteCRMHelper.SetNameValuePair("first_name", oItem.FirstName);
                    data[15] = clsSuiteCRMHelper.SetNameValuePair("account_name", oItem.CompanyName);
                    data[16] = clsSuiteCRMHelper.SetNameValuePair("salutation", oItem.Title);

                    if (sID == "")
                        data[17] = clsSuiteCRMHelper.SetNameValuePair("assigned_user_id", clsSuiteCRMHelper.GetUserId());
                    else
                        data[17] = clsSuiteCRMHelper.SetNameValuePair("id", sID);

                    _result = clsSuiteCRMHelper.SetEntry(data, "Contacts");
                    Outlook.UserProperty oProp = oItem.UserProperties["SOModifiedDate"];
                    if (oProp == null)
                        oProp = oItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                    oProp.Value = "Fresh";
                    Outlook.UserProperty oProp2 = oItem.UserProperties["SEntryID"];
                    if (oProp2 == null)
                        oProp2 = oItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                    oProp2.Value = _result;
                    oItem.Save();
                    var sItem = lContactItems.Where(a => a.oItem.EntryID == oItem.EntryID).FirstOrDefault();
                    if (sItem != default(cContactItem))
                    {
                        sItem.oItem = oItem;
                        sItem.OModifiedDate = "Fresh";
                        sItem.SEntryID = _result;
                    }
                    else
                        lContactItems.Add(new cContactItem { SEntryID = _result, OModifiedDate = "Fresh", oItem = oItem });
                }
                catch
                { }
            }
        }
        void CItems_ItemRemove()
        {
            if (!SyncInProgress && IsContactView)
            {
                if (sDelContactId != "")
                {
                    try
                    {
                        eNameValue[] data = new eNameValue[2];
                        data[0] = clsSuiteCRMHelper.SetNameValuePair("id", sDelContactId);
                        data[1] = clsSuiteCRMHelper.SetNameValuePair("deleted", "1");
                        clsSuiteCRMHelper.SetEntry(data, "Contacts");
                        lContactItems.RemoveAll(a => a.SEntryID == sDelContactId);
                    }
                    catch
                    { }
                }
                sDelContactId = "";
            }
        }
        void CoCalendarEvents_BeforeItemMove(object Item, Outlook.MAPIFolder MoveTo, ref bool Cancel)
        {
            if (!SyncInProgress && IsContactView)
            {
                sDelContactId = "";
                Outlook.ContactItem oContact = Item as Outlook.ContactItem;
                if (oContact.UserProperties != null)
                {
                    Outlook.UserProperty oProp = oContact.UserProperties["SEntryID"];
                    if (oProp != null)
                    {
                        sDelContactId = oProp.Value.ToString();
                    }
                }
            }
        }
        private void StartTaskSync()
        {
            Outlook.NameSpace oNS = this.Application.GetNamespace("mapi");
            Outlook.MAPIFolder taskFolder = GetDefaultFolder("tasks");
            Outlook.Items items = taskFolder.Items;

            items.ItemAdd -= TItems_ItemAdd;
            items.ItemChange -= TItems_ItemChange;
            items.ItemRemove -= TItems_ItemRemove;
            items.ItemAdd += TItems_ItemAdd;
            items.ItemChange += TItems_ItemChange;
            items.ItemRemove += TItems_ItemRemove;

            Outlook.MAPIFolderEvents_12_Event oCalendarEvents;
            oCalendarEvents = taskFolder as Outlook.MAPIFolderEvents_12_Event;
            if (oCalendarEvents != null)
            {
                oCalendarEvents.BeforeItemMove -= ToCalendarEvents_BeforeItemMove;
                oCalendarEvents.BeforeItemMove += ToCalendarEvents_BeforeItemMove;
            }

            SyncInProgress = true;
            GetOutlookTItems(taskFolder);
            SyncTasks(taskFolder);            
            SyncInProgress = false;

        }
        private Outlook.OlImportance GetImportance(string sImportance)
        {
            Outlook.OlImportance oPriority = Outlook.OlImportance.olImportanceLow;
            switch (sImportance)
            {
                case "High":
                    oPriority = Outlook.OlImportance.olImportanceHigh;
                    break;
                case "Medium":
                    oPriority = Outlook.OlImportance.olImportanceNormal;
                    break;
            }
            return oPriority;
        }
        private Outlook.OlTaskStatus GetStatus(string sStatus)
        {
            Outlook.OlTaskStatus oStatus = Outlook.OlTaskStatus.olTaskNotStarted;
            switch (sStatus)
            {
                case "In Progress":
                    oStatus = Outlook.OlTaskStatus.olTaskInProgress;
                    break;
                case "Completed":
                    oStatus = Outlook.OlTaskStatus.olTaskComplete;
                    break;
                case "Deferred":
                    oStatus = Outlook.OlTaskStatus.olTaskDeferred;
                    break;

            }
            return oStatus;
        }
        private void SyncTasks(Outlook.MAPIFolder tasksFolder)
        {
            eGetEntryListResult _result2 = clsSuiteCRMHelper.GetEntryList("Tasks", "",
                            0, "date_entered DESC", 0, false, clsSuiteCRMHelper.GetSugarFields("Tasks"));
            if (_result2 != null)
            {
                foreach (var oResult in _result2.entry_list)
                {
                    try
                    {
                        dynamic dResult = JsonConvert.DeserializeObject(oResult.name_value_object.ToString());
                        var oItem = lTaskItems.Where(a => a.SEntryID == dResult.id.value.ToString()).FirstOrDefault();
                        if (oItem == default(cTaskItem))
                        {
                            Outlook.TaskItem tItem = tasksFolder.Items.Add(Outlook.OlItemType.olTaskItem);
                            tItem.Subject = dResult.name.value.ToString();
                            tItem.Body = dResult.description.value.ToString();
                            if (!string.IsNullOrWhiteSpace(dResult.date_start.value.ToString()))
                            {
                                tItem.StartDate = DateTime.Parse(dResult.date_start.value.ToString());
                            }
                            if (!string.IsNullOrWhiteSpace(dResult.date_due.value.ToString()))
                            {
                                tItem.DueDate = DateTime.Parse(dResult.date_due.value.ToString());
                            }

                            tItem.Status = GetStatus(dResult.status.value.ToString());
                            tItem.Importance = GetImportance(dResult.priority.value.ToString());

                            Outlook.UserProperty oProp = tItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                            oProp.Value = dResult.date_modified.value.ToString();
                            Outlook.UserProperty oProp2 = tItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                            oProp2.Value = dResult.id.value.ToString();
                            tItem.Save();
                            lTaskItems.Add(new cTaskItem
                            {
                                oItem = tItem,
                                OModifiedDate = dResult.date_modified.value.ToString(),
                                SEntryID = dResult.id.value.ToString(),
                                Touched = true
                            });
                        }
                        else
                        {
                            oItem.Touched = true;
                            Outlook.TaskItem tItem = oItem.oItem;
                            Outlook.UserProperty oProp = tItem.UserProperties["SOModifiedDate"];

                            if (oProp.Value != dResult.date_modified.value.ToString())
                            {
                                tItem.Subject = dResult.name.value.ToString();
                                tItem.Body = dResult.description.value.ToString();
                                if (!string.IsNullOrWhiteSpace(dResult.date_start.value.ToString()))
                                {
                                    tItem.StartDate = DateTime.Parse(dResult.date_start.value.ToString());
                                }
                                if (!string.IsNullOrWhiteSpace(dResult.date_due.value.ToString()))
                                {
                                    tItem.DueDate = DateTime.Parse(dResult.date_due.value.ToString());
                                }

                                tItem.Status = GetStatus(dResult.status.value.ToString());
                                tItem.Importance = GetImportance(dResult.priority.value.ToString());
                                if (oProp == null)
                                    oProp = tItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                                oProp.Value = dResult.date_modified.value.ToString();
                                Outlook.UserProperty oProp2 = tItem.UserProperties["SEntryID"];
                                if (oProp2 == null)
                                    oProp2 = tItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                                oProp2.Value = dResult.id.value.ToString();
                                tItem.Save();
                            }
                        }
                    }
                    catch
                    { }
                }
            }

            try
            {
                var lItemToBeDeletedO = lTaskItems.Where(a => !a.Touched && !string.IsNullOrWhiteSpace(a.OModifiedDate)).ToList();
                foreach (var oItem in lItemToBeDeletedO)
                {
                    oItem.oItem.Delete();
                }
                lTaskItems.RemoveAll(a => !a.Touched && !string.IsNullOrWhiteSpace(a.OModifiedDate));

                var lItemToBeAddedToS = lTaskItems.Where(a => !a.Touched && string.IsNullOrWhiteSpace(a.OModifiedDate)).ToList();
                foreach (var oItem in lItemToBeAddedToS)
                {
                    AddTaskToS(oItem.oItem);
                }
            }
            catch
            { }
        }
        private void GetOutlookTItems(Outlook.MAPIFolder taskFolder)
        {
            if (lTaskItems == null)
            {
                lTaskItems = new List<cTaskItem>();
                Outlook.Items items = taskFolder.Items.Restrict("[MessageClass] = 'IPM.Task'");
                foreach (Outlook.TaskItem oItem in items)
                {
                    Outlook.UserProperty oProp = oItem.UserProperties["SOModifiedDate"];
                    if (oProp != null)
                    {
                        Outlook.UserProperty oProp2 = oItem.UserProperties["SEntryID"];
                        lTaskItems.Add(new cTaskItem
                        {
                            oItem = oItem,
                            OModifiedDate = oProp.Value.ToString(),
                            SEntryID = oProp2.Value.ToString()
                        });
                    }
                    else
                    {
                        lTaskItems.Add(new cTaskItem
                        {
                            oItem = oItem
                        });
                    }
                }
            }
        }

        void TItems_ItemChange(object Item)
        {
            if (!SyncInProgress && IsTaskView)
            {
                bool SyncOldStatus = SyncInProgress;
                SyncInProgress = true;
                var oItem = Item as Outlook.TaskItem;
                if (PrevTaskID != oItem.EntryID)
                {
                    Outlook.UserProperty oProp1 = oItem.UserProperties["SEntryID"];
                    if (oProp1 != null)
                    {
                        AddTaskToS(oItem, oProp1.Value.ToString());
                    }
                }                
                SyncInProgress = SyncOldStatus;
            }
        }

        void TItems_ItemAdd(object Item)
        {
            if (!SyncInProgress && IsTaskView)
            {
                AddTaskToS(Item as Outlook.TaskItem);
            }
        }
        private void AddTaskToS(Outlook.TaskItem oItem, string sID = "")
        {
            if (oItem != null)
            {
                try
                {
                    PrevTaskID = oItem.EntryID;
                    string _result = "";
                    eNameValue[] data = new eNameValue[7];
                    string strStatus = "";
                    string strImportance = "";
                    switch (oItem.Status)
                    {
                        case Outlook.OlTaskStatus.olTaskNotStarted:
                            strStatus = "Not Started";
                            break;
                        case Outlook.OlTaskStatus.olTaskInProgress:
                            strStatus = "In Progress";
                            break;
                        case Outlook.OlTaskStatus.olTaskComplete:
                            strStatus = "Completed";
                            break;
                        case Outlook.OlTaskStatus.olTaskDeferred:
                            strStatus = "Deferred";
                            break;
                    }
                    switch (oItem.Importance)
                    {
                        case Outlook.OlImportance.olImportanceLow:
                            strImportance = "Low";
                            break;

                        case Outlook.OlImportance.olImportanceNormal:
                            strImportance = "Medium";
                            break;

                        case Outlook.OlImportance.olImportanceHigh:
                            strImportance = "High";
                            break;
                    }

                    DateTime uTCDateTime = new DateTime();
                    DateTime time2 = new DateTime();
                    uTCDateTime = this.GetUTCDateTime(oItem.StartDate);
                    time2 = this.GetUTCDateTime(oItem.DueDate);
                    string str = string.Format("{0:yyyy-MM-dd HH:mm:ss}", uTCDateTime);
                    string str2 = string.Format("{0:yyyy-MM-dd HH:mm:ss}", time2);

                    data[0] = clsSuiteCRMHelper.SetNameValuePair("name", oItem.Subject);
                    data[1] = clsSuiteCRMHelper.SetNameValuePair("description", oItem.Body);
                    data[2] = clsSuiteCRMHelper.SetNameValuePair("status", strStatus);
                    data[3] = clsSuiteCRMHelper.SetNameValuePair("date_due", str2);
                    data[4] = clsSuiteCRMHelper.SetNameValuePair("date_start", str);
                    data[5] = clsSuiteCRMHelper.SetNameValuePair("priority", strImportance);
                    if (sID == "")
                        data[6] = clsSuiteCRMHelper.SetNameValuePair("assigned_user_id", clsSuiteCRMHelper.GetUserId());
                    else
                        data[6] = clsSuiteCRMHelper.SetNameValuePair("id", sID);

                    _result = clsSuiteCRMHelper.SetEntry(data, "Tasks");
                    Outlook.UserProperty oProp = oItem.UserProperties["SOModifiedDate"];
                    if (oProp == null)
                        oProp = oItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                    oProp.Value = "Fresh";
                    Outlook.UserProperty oProp2 = oItem.UserProperties["SEntryID"];
                    if (oProp2 == null)
                        oProp2 = oItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                    oProp2.Value = _result;
                    oItem.Save();
                    var sItem = lTaskItems.Where(a => a.oItem.EntryID == oItem.EntryID).FirstOrDefault();
                    if (sItem != default(cTaskItem))
                    {
                        sItem.oItem = oItem;
                        sItem.OModifiedDate = "Fresh";
                        sItem.SEntryID = _result;
                    }
                    else
                        lTaskItems.Add(new cTaskItem { SEntryID = _result, OModifiedDate = "Fresh", oItem = oItem });
                }
                catch
                { }
            }
        }
        void TItems_ItemRemove()
        {
            if (!SyncInProgress && IsTaskView)
            {
                if (sDelTaskId != "")
                {
                    try
                    {
                        eNameValue[] data = new eNameValue[2];
                        data[0] = clsSuiteCRMHelper.SetNameValuePair("id", sDelTaskId);
                        data[1] = clsSuiteCRMHelper.SetNameValuePair("deleted", "1");
                        clsSuiteCRMHelper.SetEntry(data, "Tasks");
                        lTaskItems.RemoveAll(a => a.SEntryID == sDelTaskId);
                    }
                    catch
                    { }
                }
                sDelTaskId = "";
                
            }
        }
        void ToCalendarEvents_BeforeItemMove(object Item, Outlook.MAPIFolder MoveTo, ref bool Cancel)
        {
            if (!SyncInProgress && IsTaskView)
            {
                sDelTaskId = "";
                Outlook.TaskItem oTask = Item as Outlook.TaskItem;
                if (oTask.UserProperties != null)
                {
                    Outlook.UserProperty oProp = oTask.UserProperties["SEntryID"];
                    if (oProp != null)
                    {
                        sDelTaskId = oProp.Value.ToString();                        
                    }
                }
            }
        }
        private void StartCalendarSync()
        {
            Outlook.NameSpace oNS = this.Application.GetNamespace("mapi");
            Outlook.MAPIFolder appointmentsFolder = GetDefaultFolder("appointments");
            Outlook.Items items = appointmentsFolder.Items;
            
            items.ItemAdd -= Items_ItemAdd;
            items.ItemChange -= Items_ItemChange;
            items.ItemRemove -= Items_ItemRemove;
            items.ItemAdd += Items_ItemAdd;
            items.ItemChange += Items_ItemChange;
            items.ItemRemove += Items_ItemRemove;

            Outlook.MAPIFolderEvents_12_Event oCalendarEvents;
            oCalendarEvents = appointmentsFolder as Outlook.MAPIFolderEvents_12_Event;
            if (oCalendarEvents != null)
            {
                oCalendarEvents.BeforeItemMove -= oCalendarEvents_BeforeItemMove;
                oCalendarEvents.BeforeItemMove += oCalendarEvents_BeforeItemMove;
            }

            SyncInProgress = true;
            GetOutlookCalItems(appointmentsFolder);
            SyncMeetings(appointmentsFolder, "Meetings");
            SyncMeetings(appointmentsFolder, "Calls");
            SyncInProgress = false;
                        
        }

        void oCalendarEvents_BeforeItemMove(object Item, Outlook.MAPIFolder MoveTo, ref bool Cancel)
        {
            if (!SyncInProgress && IsCalendarView)
            {
                sDelCalId = "";
                sDelCalModule = "";
                Outlook.AppointmentItem oApp = Item as Outlook.AppointmentItem;
                if (oApp.UserProperties != null)
                {
                    Outlook.UserProperty oProp = oApp.UserProperties["SEntryID"];
                    Outlook.UserProperty oProp1 = oApp.UserProperties["SType"];
                    if (oProp != null && oProp1 != null)
                    {
                        sDelCalId = oProp.Value.ToString();
                        sDelCalModule = oProp1.Value.ToString();
                    }
                }                
            }
        }

        void Items_ItemRemove()
        {
            if (!SyncInProgress && IsCalendarView)
            {
                if (sDelCalId != "")
                {
                    try
                    {
                        eNameValue[] data = new eNameValue[2];
                        data[0] = clsSuiteCRMHelper.SetNameValuePair("id", sDelCalId);
                        data[1] = clsSuiteCRMHelper.SetNameValuePair("deleted", "1");
                        clsSuiteCRMHelper.SetEntry(data, sDelCalModule);
                        lCalItems.RemoveAll(a => a.SEntryID == sDelCalId);
                    }
                    catch
                    { }
                }
                sDelCalId = "";
                sDelCalModule = "";
            }
        }

        void Items_ItemChange(object Item)
        {
            if (!SyncInProgress && IsCalendarView)
            {
                bool SyncOldStatus = SyncInProgress;
                SyncInProgress = true;
                var aItem = Item as Outlook.AppointmentItem;
                if (PrevCalSID != aItem.EntryID)
                {
                    Outlook.UserProperty oProp = aItem.UserProperties["SType"];
                    Outlook.UserProperty oProp1 = aItem.UserProperties["SEntryID"];
                    if (oProp != null && oProp1 != null)
                    {
                        AddAppointmentToS(aItem, oProp.Value.ToString(), oProp1.Value.ToString());
                    }
                }                
                SyncInProgress = SyncOldStatus;
            }
        }

        void Items_ItemAdd(object Item)
        {
            if (!SyncInProgress && IsCalendarView)
            {
                AddAppointmentToS(Item as Outlook.AppointmentItem, "Meetings");
            }
        }

        private void GetOutlookCalItems(Outlook.MAPIFolder appointmentsFolder)
        {
            if (lCalItems == null)
            {
                lCalItems = new List<cAppItem>();
                Outlook.Items items = appointmentsFolder.Items.Restrict("[MessageClass] = 'IPM.Appointment'");
                foreach (Outlook.AppointmentItem aItem in items)
                {
                    Outlook.UserProperty oProp = aItem.UserProperties["SOModifiedDate"];
                    if (oProp != null)
                    {
                        Outlook.UserProperty oProp1 = aItem.UserProperties["SType"];
                        Outlook.UserProperty oProp2 = aItem.UserProperties["SEntryID"];
                        lCalItems.Add(new cAppItem
                        {
                            oItem = aItem,
                            OModifiedDate = oProp.Value.ToString(),
                            SType = oProp1.Value.ToString(),
                            SEntryID = oProp2.Value.ToString()
                        });
                    }
                    else
                    {
                        lCalItems.Add(new cAppItem
                        {
                            oItem = aItem,
                            SType = "Meetings"
                        });
                    }
                }
            }
        }
        private void SyncMeetings(Outlook.MAPIFolder appointmentsFolder, string sModule)
        {
            eGetEntryListResult _result2 = clsSuiteCRMHelper.GetEntryList(sModule, "",
                            0, "date_entered DESC", 0, false, clsSuiteCRMHelper.GetSugarFields(sModule));
            if (_result2 != null)
            {
                foreach (var oResult in _result2.entry_list)
                {
                    try
                    {
                        dynamic dResult = JsonConvert.DeserializeObject(oResult.name_value_object.ToString());
                        var oItem = lCalItems.Where(a => a.SEntryID == dResult.id.value.ToString() && a.SType == sModule).FirstOrDefault();
                        if (oItem == default(cAppItem))
                        {
                            Outlook.AppointmentItem aItem = appointmentsFolder.Items.Add(Outlook.OlItemType.olAppointmentItem);
                            aItem.Subject = dResult.name.value.ToString();
                            aItem.Body = dResult.description.value.ToString();
                            if (!string.IsNullOrWhiteSpace(dResult.date_start.value.ToString()))
                            {
                                aItem.Start = DateTime.Parse(dResult.date_start.value.ToString());
                                int iMin = 0, iHour = 0;
                                if (!string.IsNullOrWhiteSpace(dResult.duration_minutes.value.ToString()))
                                {
                                    iMin = int.Parse(dResult.duration_minutes.value.ToString());
                                }
                                if (!string.IsNullOrWhiteSpace(dResult.duration_hours.value.ToString()))
                                {
                                    iHour = int.Parse(dResult.duration_hours.value.ToString());
                                }
                                if (sModule == "Meetings")
                                {
                                    aItem.Location = dResult.location.value.ToString();
                                    aItem.End = aItem.Start;
                                    if (iHour > 0)
                                        aItem.End.AddHours(iHour);
                                    if (iMin > 0)
                                        aItem.End.AddMinutes(iMin);
                                }
                                try
                                {
                                    aItem.Duration = iMin + iHour * 60;
                                }
                                catch (Exception)
                                { }
                            }
                            Outlook.UserProperty oProp = aItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                            oProp.Value = dResult.date_modified.value.ToString();
                            Outlook.UserProperty oProp1 = aItem.UserProperties.Add("SType", Outlook.OlUserPropertyType.olText);
                            oProp1.Value = sModule;
                            Outlook.UserProperty oProp2 = aItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                            oProp2.Value = dResult.id.value.ToString();
                            aItem.Save();
                            lCalItems.Add(new cAppItem
                            {
                                oItem = aItem,
                                OModifiedDate = dResult.date_modified.value.ToString(),
                                SType = sModule,
                                SEntryID = dResult.id.value.ToString(),
                                Touched = true
                            });
                        }
                        else
                        {
                            oItem.Touched = true;
                            Outlook.AppointmentItem aItem = oItem.oItem;
                            Outlook.UserProperty oProp = aItem.UserProperties["SOModifiedDate"];

                            if (oProp.Value != dResult.date_modified.value.ToString())
                            {
                                aItem.Subject = dResult.name.value.ToString();
                                aItem.Body = dResult.description.value.ToString();
                                if (!string.IsNullOrWhiteSpace(dResult.date_start.value.ToString()))
                                {
                                    aItem.Start = DateTime.Parse(dResult.date_start.value.ToString());
                                    int iMin = 0, iHour = 0;
                                    if (!string.IsNullOrWhiteSpace(dResult.duration_minutes.value.ToString()))
                                    {
                                        iMin = int.Parse(dResult.duration_minutes.value.ToString());
                                    }
                                    if (!string.IsNullOrWhiteSpace(dResult.duration_hours.value.ToString()))
                                    {
                                        iHour = int.Parse(dResult.duration_hours.value.ToString());
                                    }
                                    if (sModule == "Meetings")
                                    {
                                        aItem.Location = dResult.location.value.ToString();
                                        aItem.End = aItem.Start;
                                        if (iHour > 0)
                                            aItem.End.AddHours(iHour);
                                        if (iMin > 0)
                                            aItem.End.AddMinutes(iMin);
                                    }
                                    try
                                    {
                                        aItem.Duration = iMin + iHour * 60;
                                    }
                                    catch (Exception)
                                    { }
                                }
                                if (oProp == null)
                                    oProp = aItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                                oProp.Value = dResult.date_modified.value.ToString();
                                Outlook.UserProperty oProp1 = aItem.UserProperties["SType"];
                                if (oProp1 == null)
                                    oProp1 = aItem.UserProperties.Add("SType", Outlook.OlUserPropertyType.olText);
                                oProp1.Value = sModule;
                                Outlook.UserProperty oProp2 = aItem.UserProperties["SEntryID"];
                                if (oProp2 == null)
                                    oProp2 = aItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                                oProp2.Value = dResult.id.value.ToString();
                                aItem.Save();
                            }
                        }
                    }
                    catch
                    { }
                }
            }
            try
            {
                if (sModule == "Meetings")
                {
                    var lItemToBeDeletedO = lCalItems.Where(a => !a.Touched && !string.IsNullOrWhiteSpace(a.OModifiedDate) && a.SType == sModule).ToList();
                    foreach (var oItem in lItemToBeDeletedO)
                    {
                        oItem.oItem.Delete();
                    }
                    lCalItems.RemoveAll(a => !a.Touched && !string.IsNullOrWhiteSpace(a.OModifiedDate) && a.SType == sModule);
                }
                var lItemToBeAddedToS = lCalItems.Where(a => !a.Touched && string.IsNullOrWhiteSpace(a.OModifiedDate) && a.SType == sModule).ToList();
                foreach (var oItem in lItemToBeAddedToS)
                {
                    AddAppointmentToS(oItem.oItem, sModule);
                }
            }
            catch
            { }
        }
        private void AddAppointmentToS(Outlook.AppointmentItem aItem, string sModule, string sID = "")
        {
            if (aItem != null)
            {
                try
                {
                    PrevCalSID = aItem.EntryID;
                    string _result = "";
                    eNameValue[] data = new eNameValue[8];
                    DateTime uTCDateTime = new DateTime();
                    DateTime time2 = new DateTime();
                    uTCDateTime = this.GetUTCDateTime(aItem.Start);
                    time2 = this.GetUTCDateTime(aItem.End);
                    string str = string.Format("{0:yyyy-MM-dd HH:mm:ss}", uTCDateTime);
                    string str2 = string.Format("{0:yyyy-MM-dd HH:mm:ss}", time2);
                    int num = aItem.Duration / 60;
                    int num2 = aItem.Duration % 60;
                    data[0] = clsSuiteCRMHelper.SetNameValuePair("name", aItem.Subject);
                    data[1] = clsSuiteCRMHelper.SetNameValuePair("description", aItem.Body);
                    data[2] = clsSuiteCRMHelper.SetNameValuePair("location", aItem.Location);
                    data[3] = clsSuiteCRMHelper.SetNameValuePair("date_start", str);
                    data[4] = clsSuiteCRMHelper.SetNameValuePair("date_end", str2);
                    data[5] = clsSuiteCRMHelper.SetNameValuePair("duration_minutes", num2.ToString());
                    data[6] = clsSuiteCRMHelper.SetNameValuePair("duration_hours", num.ToString());
                    if (sID == "")
                        data[7] = clsSuiteCRMHelper.SetNameValuePair("assigned_user_id", clsSuiteCRMHelper.GetUserId());
                    else
                        data[7] = clsSuiteCRMHelper.SetNameValuePair("id", sID);

                    _result = clsSuiteCRMHelper.SetEntry(data, sModule);
                    Outlook.UserProperty oProp = aItem.UserProperties["SOModifiedDate"];
                    if (oProp == null)
                        oProp = aItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                    oProp.Value = "Fresh";
                    Outlook.UserProperty oProp1 = aItem.UserProperties["SType"];
                    if (oProp1 == null)
                        oProp1 = aItem.UserProperties.Add("SType", Outlook.OlUserPropertyType.olText);
                    oProp1.Value = sModule;
                    Outlook.UserProperty oProp2 = aItem.UserProperties["SEntryID"];
                    if (oProp2 == null)
                        oProp2 = aItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                    oProp2.Value = _result;
                    aItem.Save();
                    var sItem = lCalItems.Where(a => a.oItem.EntryID == aItem.EntryID).FirstOrDefault();
                    if (sItem != default(cAppItem))
                    {
                        sItem.oItem = aItem;
                        sItem.OModifiedDate = "Fresh";
                        sItem.SEntryID = _result;
                    }
                    else
                    {
                        lCalItems.Add(new cAppItem { SEntryID = _result, SType = sModule, OModifiedDate = "Fresh", oItem = aItem });
                    }
                }
                catch
                { }
            }
        }
        public DateTime GetUTCDateTime(DateTime dateTime)
        {
            return dateTime.ToUniversalTime();
        }

        
        private void cbtnArchive_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ManualArchive();
        }

        private void cbtnSettings_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            frmSettings objacbbSettings = new frmSettings();
            objacbbSettings.ShowDialog();
        }

        private void ManualArchive()
        {
            if (Globals.ThisAddIn.SuiteCRMUserSession.id == "")
            {
                frmSettings objacbbSettings = new frmSettings();
                objacbbSettings.ShowDialog();
            }
            if (Globals.ThisAddIn.SuiteCRMUserSession.id != "")
            {
                frmArchive objForm = new frmArchive();
                objForm.ShowDialog();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                if (SuiteCRMUserSession != null)
                    SuiteCRMUserSession.LogOut();
                if (this.CommandBarExists("SuiteCRM"))
                {
                    this.objSuiteCRMMenuBar2007.Delete();
                }
                this.UnregisterEvents();
            }
            catch (Exception ex)
            {
                ex.Data.Clear();
            }
        }

        private void UnregisterEvents()
        {
            try
            {
                this.Application.ItemContextMenuDisplay -= new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(this.Application_ItemContextMenuDisplay);
            }
            catch (System.Exception ex)
            {

            }

            try
            {
                this.btnArvive.Click -= new Office._CommandBarButtonEvents_ClickEventHandler(this.cbtnArchive_Click);
            }
            catch (System.Exception ex1)
            {

            }

            try
            {
                this.objExplorer.Application.NewMailEx -= new Outlook.ApplicationEvents_11_NewMailExEventHandler(this.Application_NewMail);
            }
            catch (System.Exception ex2)
            {

            }

            try
            {
                this.objExplorer.Application.ItemSend -= new Outlook.ApplicationEvents_11_ItemSendEventHandler(this.Application_ItemSend);
            }
            catch (System.Exception ex3)
            {

            }


        }

        private bool CommandBarExists(string name)
        {
            try
            {
                string text1 = Globals.ThisAddIn.Application.ActiveExplorer().CommandBars[name].Name;
                return true;
            }
            catch (System.Exception)
            {
                return false;
            }
        }

        #region VSTO generated code

        private void Application_ItemContextMenuDisplay(Office.CommandBar CommandBar, Outlook.Selection Selection)
        {
            Outlook.Selection selection = Selection;
            Outlook.MailItem item1 = (Outlook.MailItem)selection[1];
            Office.CommandBarButton objMainMenu = (Office.CommandBarButton)CommandBar.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, this.missing, this.missing, this.missing, this.missing);
            objMainMenu.Caption = "SuiteCRM Archive";
            objMainMenu.Visible = true;
            objMainMenu.Picture = RibbonImageHelper.Convert(Resources.SuiteCRM1);
            objMainMenu.Click += new Office._CommandBarButtonEvents_ClickEventHandler(this.contextMenuArchiveButton_Click);
        }

        private void contextMenuArchiveButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (Globals.ThisAddIn.SuiteCRMUserSession.id == "")
            {
                frmSettings objacbbSettings = new frmSettings();
                objacbbSettings.ShowDialog();
            }
            frmArchive objForm = new frmArchive();
            objForm.ShowDialog();
        }

        private void Application_ItemSend(object item, ref bool target)
        {
            try
            {
                if (settings.AutoArchive)
                {
                    if (item is Outlook.MailItem)
                    {
                        Outlook.MailItem objMail = (Outlook.MailItem)item;
                        if (objMail.UserProperties["SuiteCRM"] == null)
                        {
                            ArchiveEmail(objMail, 3, this.settings.ExcludedEmails);
                            objMail.UserProperties.Add("SuiteCRM", Outlook.OlUserPropertyType.olText, true, Outlook.OlUserPropertyType.olText);
                            objMail.UserProperties["SuiteCRM"].Value = "True";
                            objMail.Save();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "OutlookEvents_ItemSend General Exception:" + "\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "Data:" + ex.Data.ToString() + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------" + "\n";
                clsSuiteCRMHelper.WriteLog(strLog);
            }
        }

        private void Application_NewMail(string EntryID)
        {
            try
            {
                if (settings.AutoArchive)
                {
                    var objItem = Globals.ThisAddIn.Application.Session.GetItemFromID(EntryID);
                    if (objItem is Outlook.MailItem)
                    {
                        Outlook.MailItem objMail = (Outlook.MailItem)objItem;
                        if (objMail.UserProperties["SuiteCRM"] == null)
                        {
                            ArchiveEmail(objMail, 2, this.settings.ExcludedEmails);
                            objMail.UserProperties.Add("SuiteCRM", Outlook.OlUserPropertyType.olText, true, Outlook.OlUserPropertyType.olText);
                            objMail.UserProperties["SuiteCRM"].Value = "True";
                            objMail.Save();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "Application_NewMail General Exception:" + "\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "Data:" + ex.Data.ToString() + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------" + "\n";
                clsSuiteCRMHelper.WriteLog(strLog);
            }
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new SuiteCRMRibbon();
        }

        public void SuiteCRMAuthenticate()
        {
            if (Globals.ThisAddIn.SuiteCRMUserSession == null)
            {
                Authenticate();
            }
            else
            {
                if (Globals.ThisAddIn.SuiteCRMUserSession.id == "")
                    Authenticate();
            }

        }

        public void Authenticate()
        {
            try
            {
                string strUsername = Globals.ThisAddIn.settings.username;
                string strPassword = Globals.ThisAddIn.settings.password;                

                Globals.ThisAddIn.SuiteCRMUserSession = new SuiteCRMClient.clsUsersession("", "", "","");
                string strURL = Globals.ThisAddIn.settings.host;
                if (strURL != "")
                {
                    Globals.ThisAddIn.SuiteCRMUserSession = new SuiteCRMClient.clsUsersession(strURL, strUsername, strPassword, Globals.ThisAddIn.settings.LDAPKey);
                    Globals.ThisAddIn.SuiteCRMUserSession.AwaitingAuthentication = true;
                    try
                    {
                        if (settings.IsLDAPAuthentication)
                        {
                            Globals.ThisAddIn.SuiteCRMUserSession.AuthenticateLDAP();
                        }
                        else
                            Globals.ThisAddIn.SuiteCRMUserSession.Login();
                        if (Globals.ThisAddIn.SuiteCRMUserSession.id != "")
                            return;
                    }
                    catch (Exception ex)
                    {
                        ex.Data.Clear();
                    }
                }
                Globals.ThisAddIn.SuiteCRMUserSession.AwaitingAuthentication = false;
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "Authenticate method General Exception:\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------\n";
                clsSuiteCRMHelper.WriteLog(strLog);
                ex.Data.Clear();
            }
        }

        private void GetMailFolders(Outlook.Folders objInpFolders)
        {
            try
            {
                foreach (Outlook.Folder objFolder in objInpFolders)
                {
                    if (objFolder.Folders.Count > 0)
                    {
                        lstOutlookFolders.Add(objFolder);
                        GetMailFolders(objFolder.Folders);
                    }
                    else
                        lstOutlookFolders.Add(objFolder);
                }
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "GetMailFolders method General Exception:" + "\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "Data:" + ex.Data.ToString() + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------" + "\n";
                clsSuiteCRMHelper.WriteLog(strLog);
                ex.Data.Clear();
            }
        }

        private void ArchiveEmail(Outlook.MailItem objMail, int intArchiveType, string strExcludedEmails = "")
        {
            try
            {
                SuiteCRMClient.clsEmailArchive objEmail = new SuiteCRMClient.clsEmailArchive();
                objEmail.From = objMail.SenderEmailAddress;
                objEmail.To = "";
                foreach (Outlook.Recipient objRecepient in objMail.Recipients)
                {
                    if (objEmail.To == "")
                        objEmail.To = objRecepient.Address;
                    else
                        objEmail.To += ";" + objRecepient.Address;
                }
                objEmail.Subject = objMail.Subject;
                objEmail.Body = objMail.Body;
                objEmail.HTMLBody = objMail.HTMLBody;
                objEmail.ArchiveType = intArchiveType;
                foreach (Outlook.Attachment objMailAttachments in objMail.Attachments)
                {
                    objEmail.Attachments.Add(new SuiteCRMClient.clsEmailAttachments { DisplayName = objMailAttachments.DisplayName, FileContentInBase64String = Base64Encode(objMailAttachments, objMail) });
                }
                System.Threading.Thread objThread = new System.Threading.Thread(() => ArchiveEmailThread(objEmail, intArchiveType, strExcludedEmails));
                objThread.Start();
                
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "ArchiveEmail method General Exception:" + "\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "Data:" + ex.Data.ToString() + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------" + "\n";
                clsSuiteCRMHelper.WriteLog(strLog);
                ex.Data.Clear();
            }
        }

        private void ArchiveEmailThread(SuiteCRMClient.clsEmailArchive objEmail, int intArchiveType, string strExcludedEmails = "")
        {
            try
            {
                if (SuiteCRMUserSession != null)
                {                    
                    while (SuiteCRMUserSession.AwaitingAuthentication == true)
                    {
                        System.Threading.Thread.Sleep(1000);
                    }                    
                    if (SuiteCRMUserSession.id != "")
                    {
                        objEmail.SuiteCRMUserSession = SuiteCRMUserSession;
                        objEmail.Save(strExcludedEmails);
                    }
                }
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "ArchiveEmailThread method General Exception:" + "\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "Data:" + ex.Data.ToString() + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------" + "\n";
                clsSuiteCRMHelper.WriteLog(strLog);
                ex.Data.Clear();
            }

        }

        public byte[] Base64Encode(Outlook.Attachment objMailAttachment, Outlook.MailItem objMail)
        {
            byte[] strRet = null;
            if (objMailAttachment != null)
            {
                if (System.IO.Directory.Exists(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath") == false)
                {
                    string strPath = Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath";
                    System.IO.Directory.CreateDirectory(strPath);
                }
                try
                {
                    objMailAttachment.SaveAsFile(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath\\" + objMailAttachment.FileName);
                    strRet = System.IO.File.ReadAllBytes(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath\\" + objMailAttachment.FileName);
                }
                catch (COMException ex)
                {
                    try
                    {
                        string strLog;
                        strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                        strLog += "AddInModule.Base64Encode method COM Exception:" + "\n";
                        strLog += "Message:" + ex.Message + "\n";
                        strLog += "Source:" + ex.Source + "\n";
                        strLog += "StackTrace:" + ex.StackTrace + "\n";
                        strLog += "Data:" + ex.Data.ToString() + "\n";
                        strLog += "HResult:" + ex.HResult.ToString() + "\n";
                        strLog += "Inputs:" + "\n";
                        strLog += "Data:" + objMailAttachment.DisplayName + "\n";
                        strLog += "-------------------------------------------------------------------------" + "\n";
                        clsSuiteCRMHelper.WriteLog(strLog);
                        ex.Data.Clear();
                        string strName = Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath\\" + DateTime.Now.ToString("MMddyyyyHHmmssfff") + ".html";
                        objMail.SaveAs(strName, Microsoft.Office.Interop.Outlook.OlSaveAsType.olHTML);
                        foreach (string strFileName in System.IO.Directory.GetFiles(strName.Replace(".html", "_files")))
                        {
                            if (strFileName.EndsWith("\\" + objMailAttachment.DisplayName))
                            {
                                strRet = System.IO.File.ReadAllBytes(strFileName);
                                break;
                            }
                        }
                    }
                    catch (Exception ex1)
                    {
                        string strLog;
                        strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                        strLog += "AddInModule.Base64Encode method General Exception:" + "\n";
                        strLog += "Message:" + ex.Message + "\n";
                        strLog += "Source:" + ex.Source + "\n";
                        strLog += "StackTrace:" + ex.StackTrace + "\n";
                        strLog += "Data:" + ex.Data.ToString() + "\n";
                        strLog += "HResult:" + ex.HResult.ToString() + "\n";
                        strLog += "Inputs:" + "\n";
                        strLog += "Data:" + objMailAttachment.DisplayName + "\n";
                        strLog += "-------------------------------------------------------------------------" + "\n";
                        clsSuiteCRMHelper.WriteLog(strLog);
                        ex1.Data.Clear();
                    }
                }
                finally
                {
                    if (System.IO.Directory.Exists(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath") == true)
                    {
                        System.IO.Directory.Delete(Environment.SpecialFolder.MyDocuments.ToString(), true);
                    }
                }
            }

            return strRet;
        }
        public Outlook.MAPIFolder GetDefaultFolder(string type)
        {
            switch (type)
            {
                case "contacts":
                    return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);

                case "appointments":
                    return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

                case "tasks":
                    return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks);
            }
            return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
        }
        private void ArchiveFolderItems(Outlook.Folder objFolder, DateTime? dtAutoArchiveFrom = null)
        {
            try
            {
                Outlook.Items UnReads;
                if (dtAutoArchiveFrom == null)
                    UnReads = objFolder.Items.Restrict("[Unread]=true");
                else
                    UnReads = objFolder.Items.Restrict("[ReceivedTime] >= '" + ((DateTime)dtAutoArchiveFrom).AddDays(-1).ToString("yyyy-MM-dd HH:mm") + "'");

                for (int intItr = 1; intItr <= UnReads.Count; intItr++)
                {
                    if (UnReads[intItr] is Outlook.MailItem)
                    {
                        Outlook.MailItem objMail = (Outlook.MailItem)UnReads[intItr];

                        if (objMail.UserProperties["SuiteCRM"] == null)
                        {
                            ArchiveEmail(objMail, 2, this.settings.ExcludedEmails);
                            objMail.UserProperties.Add("SuiteCRM", Outlook.OlUserPropertyType.olText, true, Outlook.OlUserPropertyType.olText);
                            objMail.UserProperties["SuiteCRM"].Value = "True";
                            objMail.Save();
                        }
                        else
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "ArchiveFolderItems method General Exception:" + "\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "Data:" + ex.Data.ToString() + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------" + "\n";
                clsSuiteCRMHelper.WriteLog(strLog);
                ex.Data.Clear();
            }
        }

        public void ProcessMails(DateTime? dtAutoArchiveFrom = null)
        {
            if (settings.AutoArchive == false)
                return;
            System.Threading.Thread.Sleep(5000);
            while (1 == 1)
            {
                try
                {
                    lstOutlookFolders = new List<Outlook.Folder>();
                    GetMailFolders(Globals.ThisAddIn.Application.Session.Folders);
                    if (lstOutlookFolders != null)
                    {
                        foreach (Outlook.Folder objFolder in lstOutlookFolders)
                        {
                            if (settings.AutoArchiveFolders == null)
                                ArchiveFolderItems(objFolder, dtAutoArchiveFrom);
                            else if (settings.AutoArchiveFolders.Count == 0)
                                ArchiveFolderItems(objFolder, dtAutoArchiveFrom);
                            else
                            {
                                if (settings.AutoArchiveFolders.Contains(objFolder.EntryID))
                                {
                                    ArchiveFolderItems(objFolder, dtAutoArchiveFrom);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    string strLog;
                    strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                    strLog += "ProcessMails method General Exception:" + "\n";
                    strLog += "Message:" + ex.Message + "\n";
                    strLog += "Source:" + ex.Source + "\n";
                    strLog += "StackTrace:" + ex.StackTrace + "\n";
                    strLog += "Data:" + ex.Data.ToString() + "\n";
                    strLog += "HResult:" + ex.HResult.ToString() + "\n";
                    strLog += "-------------------------------------------------------------------------" + "\n";
                    clsSuiteCRMHelper.WriteLog(strLog);
                    ex.Data.Clear();
                }
                if (dtAutoArchiveFrom != null)
                    break;

                System.Threading.Thread.Sleep(5000);
            }
        }        
    }
}
