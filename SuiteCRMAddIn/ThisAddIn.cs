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
        private bool IsCalendarView = false;
        
        List<cTaskItem> lTaskItems;
        private string sDelTaskId = "";
        private bool IsTaskView = false;
        
        List<cContactItem> lContactItems;
        private string sDelContactId = "";
        private bool IsContactView = false;
        
        public Office.IRibbonUI RibbonUI { get; set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
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
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.ThisAddIn_Startup");
            }
        }
              

        void objExplorer_FolderSwitch()
        {
            try
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
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.objExplorer_FolderSwitch");
            }
        }

        public async void Sync()
        //public void Sync()
        {
            try
            {
                while (true)
                {
                    if (SuiteCRMUserSession != null && SuiteCRMUserSession.id != "")
                    {
                        if (settings.SyncCalendar)
                        {
                            //StartCalendarSync();
                            // for test !!!
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
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.Sync");
            }
        }
        private void StartContactSync()
        {
            try
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
                                
                GetOutlookCItems(contactsFolder);
                SyncContacts(contactsFolder);
                
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.StartContactSync");
            }

        }
        private void SyncContacts(Outlook.MAPIFolder contactFolder)
        {
            clsSuiteCRMHelper.WriteLog("ThisAddIn.SyncContacts");
            try
            {
                int iOffset = 0;
                bool IsDone = false;
                while (true)
                {
                    bool HasAccess=false;
                    try
                    {
                        eModuleList oList = clsSuiteCRMHelper.GetModules();
                        HasAccess = oList.modules1.FirstOrDefault(a => a.module_label == "Contacts")
                            .module_acls1.FirstOrDefault(b => b.action == "export").access;
                    }
                    catch(Exception)
                    {

                    }
                    if (!HasAccess)
                        break;
                    eGetEntryListResult _result2 = clsSuiteCRMHelper.GetEntryList("Contacts", 
                                    "contacts.assigned_user_id = '" + clsSuiteCRMHelper.GetUserId() + "'",
                                    0, "date_entered DESC",iOffset, false, clsSuiteCRMHelper.GetSugarFields("Contacts"));
                    if (_result2 != null)
                    {
                        if (iOffset == _result2.next_offset)
                            break;
                        foreach (var oResult in _result2.entry_list)
                        {
                            try
                            {
                                dynamic dResult = JsonConvert.DeserializeObject(oResult.name_value_object.ToString());

                                /*
                                FOR DEBUG

                                clsSuiteCRMHelper.WriteLog("---------------------------------------");
                                clsSuiteCRMHelper.WriteLog(Convert.ToString(dResult));
                                clsSuiteCRMHelper.WriteLog("---------------------------------------");

                                clsSuiteCRMHelper.WriteLog("sync_contact = "+ dResult.sync_contact.value.ToString());
                                clsSuiteCRMHelper.WriteLog("============================");
                                */


                                var oItem = lContactItems.FirstOrDefault(a => a.SEntryID == dResult.id.value.ToString());
                                if (oItem == default(cContactItem))
                                {
                                    if (dResult.sync_contact.value.ToString() != "True")
                                    {
                                        clsSuiteCRMHelper.WriteLog("not sync!");
                                        continue;
                                    }

                                    clsSuiteCRMHelper.WriteLog("    default sync");
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
                                    lContactItems.Add(new cContactItem
                                    {
                                        oItem = cItem,
                                        OModifiedDate = DateTime.ParseExact(dResult.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null),
                                        SEntryID = dResult.id.value.ToString(),
                                        Touched = true
                                    });
                                    clsSuiteCRMHelper.WriteLog(cItem.FullName + "     is saving with " + cItem.Sensitivity.ToString());
                                    cItem.Save();                                    
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

                                        if (dResult.primary_address_street.value != null)
                                            cItem.BusinessAddressStreet = dResult.primary_address_street.value.ToString();
                                        cItem.Body = dResult.description.value.ToString();
                                        cItem.Account = cItem.CompanyName = "";
                                        if (dResult.account_name != null && dResult.account_name.value != null)
                                        {
                                            cItem.Account = dResult.account_name.value.ToString();
                                            cItem.CompanyName = dResult.account_name.value.ToString();
                                        }

                                        cItem.BusinessFaxNumber = dResult.phone_fax.value.ToString();
                                        cItem.Title = dResult.salutation.value.ToString();
                                        if (oProp == null)
                                            oProp = cItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                                        oProp.Value = dResult.date_modified.value.ToString();
                                        Outlook.UserProperty oProp2 = cItem.UserProperties["SEntryID"];
                                        if (oProp2 == null)
                                            oProp2 = cItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                                        oProp2.Value = dResult.id.value.ToString();
                                        clsSuiteCRMHelper.WriteLog("    save not default");
                                        clsSuiteCRMHelper.WriteLog(cItem.FullName+ "     is saving with" + cItem.Sensitivity.ToString());
                                        cItem.Save();
                                    }
                                    clsSuiteCRMHelper.WriteLog(cItem.FullName + " dResult.date_modified= " + dResult.date_modified.ToString());
                                    oItem.OModifiedDate = DateTime.ParseExact(dResult.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);
                                }
                            }
                            catch (Exception ex)
                            {
                                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.SyncContacts");
                            }
                        }

                    }
                    if (iOffset == _result2.next_offset)
                        iOffset = 0;
                    else
                        iOffset = _result2.next_offset;
                    if (iOffset == 0 || IsDone)
                        break;
                }
                try
                {
                    var lItemToBeDeletedO = lContactItems.Where(a => !a.Touched && a.oItem.Sensitivity == Outlook.OlSensitivity.olNormal && !string.IsNullOrWhiteSpace(a.OModifiedDate.ToString()));
                    foreach (var oItem in lItemToBeDeletedO)
                    {
                        oItem.oItem.Delete();
                    }
                    lContactItems.RemoveAll(a => !a.Touched && !string.IsNullOrWhiteSpace(a.OModifiedDate.ToString()));

                    var lItemToBeAddedToS = lContactItems.Where(a => !a.Touched && a.oItem.Sensitivity == Outlook.OlSensitivity.olNormal && string.IsNullOrWhiteSpace(a.OModifiedDate.ToString()));
                    foreach (var oItem in lItemToBeAddedToS)
                    {
                        AddContactToS(oItem.oItem);
                    }
                }
                catch (Exception ex)
                {
                    clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.SyncContacts");
                }
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.SyncContacts");
            }
        }
        private void GetOutlookCItems(Outlook.MAPIFolder taskFolder)
        {
            try
            {
                if (lContactItems == null)
                {
                    lContactItems = new List<cContactItem>();
                    Outlook.Items items = taskFolder.Items.Restrict("[MessageClass] = 'IPM.Contact'");
                    foreach (Outlook.ContactItem oItem in items)
                    {
                        //if (oItem.Sensitivity != Outlook.OlSensitivity.olPrivate)
                        //{
                            //Outlook.UserProperty sensitivityCached = oItem.UserProperties["SensitivityCached"];
                            //sensitivityCached.Value = "olNormal";
                            Outlook.UserProperty oProp = oItem.UserProperties["SOModifiedDate"];
                            if (oProp != null)
                            {
                                Outlook.UserProperty oProp2 = oItem.UserProperties["SEntryID"];
                                //clsSuiteCRMHelper.WriteLog("GetLocalContacts SOModifiedDate: " + oProp.Value.ToString());
                                DateTime modDateTime = DateTime.UtcNow;
                                if (!DateTime.TryParseExact(oProp.Value.ToString(), "yyyy-MM-dd HH:mm:ss", null, DateTimeStyles.None, out modDateTime))
                                {
                                    DateTime.TryParse(oProp.Value.ToString(), out modDateTime);
                                }
                                lContactItems.Add(new cContactItem
                                {
                                    oItem = oItem,
                                    OModifiedDate = modDateTime,
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
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.GetOutlookCItems");
            }
        }

        void CItems_ItemChange(object Item)
        {
            clsSuiteCRMHelper.WriteLog("ItemChange");

            try
            {
                var oItem = Item as Outlook.ContactItem;

                clsSuiteCRMHelper.WriteLog(oItem.FullName + " Sensitivity= " + oItem.Sensitivity);
                string entryId = oItem.EntryID;
                clsSuiteCRMHelper.WriteLog("oItem.EntryID: " + entryId);
                cContactItem contact = lContactItems.FirstOrDefault(a => a.oItem.EntryID == entryId);
                clsSuiteCRMHelper.WriteLog("EntryID=  " + oItem.EntryID);
                if (contact != default(cContactItem))
                {
                    if ((int)Math.Abs((DateTime.UtcNow - contact.OModifiedDate).TotalSeconds) > 5)
                    {
                        contact.IsUpdate = 0;
                    }

                    clsSuiteCRMHelper.WriteLog("Before UtcNow - contact.OModifiedDate= " +
                                               (int) (DateTime.UtcNow - contact.OModifiedDate).TotalSeconds);
                    clsSuiteCRMHelper.WriteLog("IsUpdate before time check: " + contact.IsUpdate.ToString());
                    if ((int) Math.Abs((DateTime.UtcNow - contact.OModifiedDate).TotalSeconds) > 2 && contact.IsUpdate == 0)
                    {
                        contact.OModifiedDate = DateTime.UtcNow;
                        clsSuiteCRMHelper.WriteLog("Change IsUpdate = " + contact.IsUpdate);
                        contact.IsUpdate++;
                    }

                    clsSuiteCRMHelper.WriteLog("contact = " + contact.oItem.FullName);
                    clsSuiteCRMHelper.WriteLog("contact mod_date= " + contact.OModifiedDate.ToString());
                    clsSuiteCRMHelper.WriteLog("UtcNow - contact.OModifiedDate= " +
                                               (int) (DateTime.UtcNow - contact.OModifiedDate).TotalSeconds);
                }
                else
                {
                    clsSuiteCRMHelper.WriteLog("not found contact. AddContactToS(oItem) ");
                }
                // oItem.Sensitivity == Outlook.OlSensitivity.olNormal
                if (IsContactView && lContactItems.Exists(a => a.oItem.EntryID == oItem.EntryID
                                                               && contact.IsUpdate == 1
                                                               && oItem.Sensitivity == Outlook.OlSensitivity.olNormal))
                {
                    Outlook.UserProperty oProp1 = oItem.UserProperties["SEntryID"];

                    if (oProp1 != null)
                    {
                        contact.IsUpdate++;
                        clsSuiteCRMHelper.WriteLog("Go to AddContactToS");
                        AddContactToS(oItem, oProp1.Value.ToString());
                    }
                    else
                    {
                        AddContactToS(oItem);
                    }
                }
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.CItems_ItemChange");
            }
            finally
            {
                clsSuiteCRMHelper.WriteLog("lContactItems.Count = " + lContactItems.Count);
            }
        }

        void CItems_ItemAdd(object Item)
        {
            try
            {
                if (!IsContactView)
                    return;

                var item = Item as Outlook.ContactItem;
                if (item.Sensitivity != Outlook.OlSensitivity.olNormal)
                {
                    lContactItems.Add(new cContactItem {OModifiedDate = DateTime.UtcNow, oItem = item });
                    clsSuiteCRMHelper.WriteLog("Contact with abnormal Sensitivity was added to lContactItems - " + item.FullName);
                    return;
                }
                    
                //Outlook.UserProperty oProp = item.UserProperties["SOModifiedDate"];
                Outlook.UserProperty oProp2 = item.UserProperties["SEntryID"];  // to avoid duplicating of the contact
                if (oProp2 != null)
                {
                    AddContactToS(item, oProp2.Value);
                }
                else 
                {
                    AddContactToS(item);
                }                
            
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.CItems_ItemAdd");
            }
        }
        private void AddContactToS(Outlook.ContactItem oItem, string sID = "")
        {
            if (!settings.SyncContacts)
                return;
            if (oItem != null && oItem.Sensitivity.ToString() == "olNormal")
            {
                try
                {
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

                    oProp.Value = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss");
                   
                    Outlook.UserProperty oProp2 = oItem.UserProperties["SEntryID"];
                    if (oProp2 == null)
                        oProp2 = oItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                    oProp2.Value = _result;

                    clsSuiteCRMHelper.WriteLog(oItem.FullName + " from save Sensitivity= " + oItem.Sensitivity);

                    if (oItem.Sensitivity.ToString() != "olNormal")
                        return;

                    clsSuiteCRMHelper.WriteLog("        Save");
                    oItem.Save();

                    string entryId = oItem.EntryID;
                    var sItem = lContactItems.FirstOrDefault(a => a.oItem.EntryID == entryId);
                    if (sItem != default(cContactItem))
                    {
                        sItem.oItem = oItem;
                        clsSuiteCRMHelper.WriteLog("ThisAddIn.AddContactToS (DateTime.UtcNow - sItem.OModifiedDate).Milliseconds = " +
                            (DateTime.UtcNow - sItem.OModifiedDate).TotalSeconds.ToString());

                        sItem.OModifiedDate = DateTime.UtcNow;

                        sItem.SEntryID = _result;
                        clsSuiteCRMHelper.WriteLog("ThisAddIn.AddContactToS sItem.OModifiedDate = "+ sItem.OModifiedDate.ToString());
                    }
                    else
                    {
                        clsSuiteCRMHelper.WriteLog("ThisAddIn.AddContactToS ADD lContactItemsFresh");
                        lContactItems.Add(new cContactItem { SEntryID = _result, OModifiedDate = DateTime.UtcNow, oItem = oItem });
                    }
                        
                }
                catch (Exception ex)
                {
                    clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.AddContactToS");
                }
            }
        }
        void CItems_ItemRemove()
        {
            if (!IsContactView) return;
            if (sDelContactId != "")
            {
                try
                {
                    foreach (var oItem in lContactItems)
                    {
                        try
                        {
                            if (oItem.oItem.Sensitivity != Outlook.OlSensitivity.olNormal)
                                continue;
                            string sID = oItem.oItem.EntryID;
                        }
                        catch (COMException ex)
                        {
                            eNameValue[] data = new eNameValue[2];
                            data[0] = clsSuiteCRMHelper.SetNameValuePair("id", oItem.SEntryID);
                            data[1] = clsSuiteCRMHelper.SetNameValuePair("deleted", "1");
                            clsSuiteCRMHelper.SetEntry(data, "Contacts");
                            oItem.Delete = true;
                        }
                    }
                    lContactItems.RemoveAll(a => a.Delete);                        
                }
                catch (Exception ex)
                {
                    clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.CItems_ItemRemove");
                }
            }
            sDelContactId = "";
        }
        
        private void StartTaskSync()
        {
            try
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
                
                GetOutlookTItems(taskFolder);
                SyncTasks(taskFolder);
                
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.StartTaskSync");
            }
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

            clsSuiteCRMHelper.WriteLog("SyncTasks");
            clsSuiteCRMHelper.WriteLog("My UserId= " + clsSuiteCRMHelper.GetUserId());
            try
            {
                int iOffset = 0;
                bool IsDone = false;
                while (true)
                {
                    eGetEntryListResult _result2 = clsSuiteCRMHelper.GetEntryList("Tasks", "",
                                    0, "date_start DESC", iOffset, false, clsSuiteCRMHelper.GetSugarFields("Tasks"));
                    if (_result2 != null)
                    {
                        if (iOffset == _result2.next_offset)
                            break;


                        foreach (var oResult in _result2.entry_list)
                        {
                            try
                            {
                                dynamic dResult = JsonConvert.DeserializeObject(oResult.name_value_object.ToString());
                                //
                                if (clsSuiteCRMHelper.GetUserId() != dResult.assigned_user_id.value.ToString())
                                    continue;

                                /*DateTime date_start = new DateTime();
                                DateTime date_due = new DateTime();*/

                                DateTime? date_start = null;
                                DateTime? date_due = null;

                                string time_start = "--:--", time_due = "--:--";


                               /* clsSuiteCRMHelper.WriteLog("---------------------------------");
                                clsSuiteCRMHelper.WriteLog("dResult= "+ Convert.ToString(dResult));
                                clsSuiteCRMHelper.WriteLog("---------------------------------");*/


                                if (!string.IsNullOrWhiteSpace(dResult.date_start.value.ToString()) && !string.IsNullOrEmpty(dResult.date_start.value.ToString()))
                                {
                                    clsSuiteCRMHelper.WriteLog("    SET date_start = dResult.date_start");
                                    date_start = DateTime.ParseExact(dResult.date_start.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);

                                    date_start = date_start.Value.Add(new DateTimeOffset(DateTime.Now).Offset);
                                    time_start = TimeSpan.FromHours(date_start.Value.Hour).Add(TimeSpan.FromMinutes(date_start.Value.Minute)).ToString(@"hh\:mm");
                                }
                                /*else
                                {
                                    clsSuiteCRMHelper.WriteLog("    SET date_start = dResult.date_modified");
                                    date_start = DateTime.Parse(dResult.date_modified.value.ToString());
                                }

                                date_start = date_start.Value.Add(new DateTimeOffset(DateTime.Now).Offset);
                                time_start = TimeSpan.FromHours(date_start.Value.Hour).Add(TimeSpan.FromMinutes(date_start.Value.Minute)).ToString(@"hh\:mm");*/

                                if (date_start != null && date_start < GetStartDate())
                                {
                                    clsSuiteCRMHelper.WriteLog("    date_start="+ date_start.ToString() + ", GetStartDate= " + GetStartDate().ToString());
                                    continue;
                                }

                                if (!string.IsNullOrWhiteSpace(dResult.date_due.value.ToString()))
                                {
                                    date_due = DateTime.ParseExact(dResult.date_due.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);
                                    date_due = date_due.Value.Add(new DateTimeOffset(DateTime.Now).Offset);
                                    time_due = TimeSpan.FromHours(date_due.Value.Hour).Add(TimeSpan.FromMinutes(date_due.Value.Minute)).ToString(@"hh\:mm");;
                                }

                                foreach (var lt in lTaskItems)
                                {
                                    clsSuiteCRMHelper.WriteLog("    Task= " + lt.SEntryID);
                                }

                                var oItem = lTaskItems.FirstOrDefault(a => a.SEntryID == dResult.id.value.ToString());


                                if (oItem == default(cTaskItem))
                                {
                                    clsSuiteCRMHelper.WriteLog("    if default");
                                    Outlook.TaskItem tItem = tasksFolder.Items.Add(Outlook.OlItemType.olTaskItem);
                                    tItem.Subject = dResult.name.value.ToString();
                                    
                                    if (!string.IsNullOrWhiteSpace(dResult.date_start.value.ToString()))
                                    {
                                        tItem.StartDate = date_start.Value;
                                    }
                                    if (!string.IsNullOrWhiteSpace(dResult.date_due.value.ToString()))
                                    {
                                        tItem.DueDate = date_due.Value;// DateTime.Parse(dResult.date_due.value.ToString());
                                    }

                                    string body = dResult.description.value.ToString();
                                    tItem.Body = string.Concat(body, "#<", time_start, "#", time_due);
                                    tItem.Status = GetStatus(dResult.status.value.ToString());
                                    tItem.Importance = GetImportance(dResult.priority.value.ToString());

                                    Outlook.UserProperty oProp = tItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                                    oProp.Value = dResult.date_modified.value.ToString();
                                    Outlook.UserProperty oProp2 = tItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                                    oProp2.Value = dResult.id.value.ToString();
                                    lTaskItems.Add(new cTaskItem
                                    {
                                        oItem = tItem,
                                        OModifiedDate = DateTime.ParseExact(dResult.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null),                                 
                                        SEntryID = dResult.id.value.ToString(),
                                        Touched = true
                                    });
                                    clsSuiteCRMHelper.WriteLog("    save 0");
                                    tItem.Save();
                                }
                                else
                                {
                                    clsSuiteCRMHelper.WriteLog("    else not default");
                                    oItem.Touched = true;
                                    Outlook.TaskItem tItem = oItem.oItem;
                                    Outlook.UserProperty oProp = tItem.UserProperties["SOModifiedDate"];

                                    clsSuiteCRMHelper.WriteLog("    oProp.Value= " + oProp.Value + ", dResult.date_modified=" + dResult.date_modified.value.ToString());
                                    if (oProp.Value != dResult.date_modified.value.ToString())
                                    {
                                        tItem.Subject = dResult.name.value.ToString();

                                        if (!string.IsNullOrWhiteSpace(dResult.date_start.value.ToString()))
                                        {
                                            clsSuiteCRMHelper.WriteLog("    tItem.StartDate= "+ tItem.StartDate+ ", date_start=" + date_start);
                                            tItem.StartDate = date_start.Value;
                                        }
                                        if (!string.IsNullOrWhiteSpace(dResult.date_due.value.ToString()))
                                        {
                                            tItem.DueDate = date_due.Value;// DateTime.Parse(dResult.date_due.value.ToString());
                                        }

                                        string body = dResult.description.value.ToString();
                                        tItem.Body = string.Concat(body, "#<", time_start, "#", time_due);
                                        tItem.Status = GetStatus(dResult.status.value.ToString());
                                        tItem.Importance = GetImportance(dResult.priority.value.ToString());
                                        if (oProp == null)
                                            oProp = tItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                                        oProp.Value = dResult.date_modified.value.ToString();
                                        Outlook.UserProperty oProp2 = tItem.UserProperties["SEntryID"];
                                        if (oProp2 == null)
                                            oProp2 = tItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                                        oProp2.Value = dResult.id.value.ToString();
                                        clsSuiteCRMHelper.WriteLog("    save 1");
                                        tItem.Save();
                                    }
                                    oItem.OModifiedDate = DateTime.ParseExact(dResult.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);
                                }
                            }
                            catch (Exception ex)
                            {
                                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.SyncTasks");
                            }
                        }
                    }
                    if (iOffset == _result2.next_offset)
                        iOffset = 0;
                    else
                        iOffset = _result2.next_offset;
                    if (iOffset == 0 || IsDone)
                        break;
                }
                try
                {
                    var lItemToBeDeletedO = lTaskItems.Where(a => !a.Touched && !string.IsNullOrWhiteSpace(a.OModifiedDate.ToString()));
                    foreach (var oItem in lItemToBeDeletedO)
                    {
                        oItem.oItem.Delete();
                    }
                    lTaskItems.RemoveAll(a => !a.Touched && !string.IsNullOrWhiteSpace(a.OModifiedDate.ToString()));

                    var lItemToBeAddedToS = lTaskItems.Where(a => !a.Touched && string.IsNullOrWhiteSpace(a.OModifiedDate.ToString()));
                    foreach (var oItem in lItemToBeAddedToS)
                    {
                        AddTaskToS(oItem.oItem);
                    }
                }
                catch (Exception ex)
                {
                    clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.SyncTasks");
                }
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.SyncTasks");
            }
        }
        private void GetOutlookTItems(Outlook.MAPIFolder taskFolder)
        {
            try
            {
                if (lTaskItems == null)
                {
                    lTaskItems = new List<cTaskItem>();
                    Outlook.Items items = taskFolder.Items; //.Restrict("[MessageClass] = 'IPM.Task'" + GetStartDateString());
                    foreach (Outlook.TaskItem oItem in items)
                    {
                        if (oItem.DueDate < DateTime.Now.AddDays(-5))
                            continue;
                        Outlook.UserProperty oProp = oItem.UserProperties["SOModifiedDate"];
                        if (oProp != null)
                        {
                            Outlook.UserProperty oProp2 = oItem.UserProperties["SEntryID"];
                            lTaskItems.Add(new cTaskItem
                            {
                                oItem = oItem,
                                //OModifiedDate = "Fresh",
                                OModifiedDate = DateTime.UtcNow,
                               
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
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.GetOutlookTItems");
            }
        }

        void TItems_ItemChange(object Item)
        {
            clsSuiteCRMHelper.WriteLog("TItems_ItemChange");
            try
            {
                var oItem = Item as Outlook.TaskItem;
                string entryId = oItem.EntryID;
                clsSuiteCRMHelper.WriteLog("    oItem.EntryID= "+ entryId);

                cTaskItem taskitem = lTaskItems.FirstOrDefault(a => a.oItem.EntryID == entryId);
                if (taskitem != default(cTaskItem))
                {
                    if ((DateTime.UtcNow - taskitem.OModifiedDate).TotalSeconds > 5)
                    {
                        clsSuiteCRMHelper.WriteLog("2 callitem.IsUpdate = " + taskitem.IsUpdate);
                        taskitem.IsUpdate = 0;
                    }

                    clsSuiteCRMHelper.WriteLog("Before UtcNow - callitem.OModifiedDate= " + (DateTime.UtcNow - taskitem.OModifiedDate).TotalSeconds.ToString());

                    if ( (int)(DateTime.UtcNow - taskitem.OModifiedDate).TotalSeconds > 2 && taskitem.IsUpdate == 0)
                    {
                        taskitem.OModifiedDate = DateTime.UtcNow;
                        clsSuiteCRMHelper.WriteLog("1 callitem.IsUpdate = " + taskitem.IsUpdate);
                        taskitem.IsUpdate++;
                    }

                    clsSuiteCRMHelper.WriteLog("callitem = " + taskitem.oItem.Subject);
                    clsSuiteCRMHelper.WriteLog("callitem.SEntryID = " + taskitem.SEntryID);
                    clsSuiteCRMHelper.WriteLog("callitem mod_date= " + taskitem.OModifiedDate.ToString());
                    clsSuiteCRMHelper.WriteLog("UtcNow - callitem.OModifiedDate= " + (DateTime.UtcNow - taskitem.OModifiedDate).TotalSeconds.ToString());
                }
                else
                {
                    clsSuiteCRMHelper.WriteLog("not found callitem ");
                }


                if (IsTaskView && lTaskItems.Exists(a => a.oItem.EntryID == entryId //// if (IsTaskView && lTaskItems.Exists(a => a.oItem.EntryID == entryId && a.OModifiedDate != "Fresh"))
                                 && taskitem.IsUpdate == 1
                                 )
                )                   
                {
                  
                    Outlook.UserProperty oProp1 = oItem.UserProperties["SEntryID"];
                    if (oProp1 != null)
                    {
                        clsSuiteCRMHelper.WriteLog("    go to AddTaskToS");
                        taskitem.IsUpdate++;
                        AddTaskToS(oItem, oProp1.Value.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.TItems_ItemChange");
            }
        }

        void TItems_ItemAdd(object Item)
        {
            try
            {
                if (IsTaskView)
                {
                    var item = Item as Outlook.TaskItem;
                    Outlook.UserProperty oProp2 = item.UserProperties["SEntryID"];  // to avoid duplicating of the task
                    if (oProp2 != null)
                    {
                        AddTaskToS(item, oProp2.Value);
                    }
                    else
                    {
                        AddTaskToS(item);
                    }
                }
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.TItems_ItemAdd");
            }
        }
        private void AddTaskToS(Outlook.TaskItem oItem, string sID = "")
        {
            clsSuiteCRMHelper.WriteLog("AddTaskToS");
            //if (!settings.SyncCalendar)
            //    return;
            if (oItem != null)
            {
                try
                {
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
                    if (oItem.DueDate != null)
                        time2 = this.GetUTCDateTime(oItem.DueDate);

                    string body = "";
                    string str, str2;
                    str = str2 = "";
                    if (oItem.Body != null)
                    {
                        body = oItem.Body.ToString();
                        var times = this.ParseTimesFromTaskBody(body);
                        if (times != null)
                        {
                            uTCDateTime = uTCDateTime.Add(times[0]);
                            time2 = time2.Add(times[1]);

                            //check max date, date must has value !
                            if (uTCDateTime.ToUniversalTime().Year < 4000)
                                str = string.Format("{0:yyyy-MM-dd HH:mm:ss}", uTCDateTime.ToUniversalTime());
                            if (time2.ToUniversalTime().Year < 4000)
                                str2 = string.Format("{0:yyyy-MM-dd HH:mm:ss}", time2.ToUniversalTime());
                        }
                        else
                        {
                            str = oItem.StartDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
                            str2 = oItem.DueDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
                        }
               
                    }
                    else
                    {
                        str = oItem.StartDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
                        str2 = oItem.DueDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
                    }

                    //str = "2016-11-10 11:34:01";
                    //str2 = "2016-11-19 11:34:01";


                    string description = "";
                    
                    if (!string.IsNullOrEmpty(body))
                    {
                        int lastIndex = body.LastIndexOf("#<");
                        if (lastIndex >= 0)
                            description = body.Remove(lastIndex);
                        else
                        {
                            description = body;
                        }
                    }
                    clsSuiteCRMHelper.WriteLog("    description= "+ description);

                    data[0] = clsSuiteCRMHelper.SetNameValuePair("name", oItem.Subject);
                    data[1] = clsSuiteCRMHelper.SetNameValuePair("description", description);
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
                    oProp.Value = DateTime.UtcNow;
                    Outlook.UserProperty oProp2 = oItem.UserProperties["SEntryID"];
                    if (oProp2 == null)
                        oProp2 = oItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                    oProp2.Value = _result;
                    string entryId = oItem.EntryID;
                    oItem.Save();

                    var sItem = lTaskItems.FirstOrDefault(a => a.oItem.EntryID == entryId);
                    if (sItem != default(cTaskItem))
                    {
                        sItem.Touched = true;
                        sItem.oItem = oItem;
                        sItem.OModifiedDate = DateTime.UtcNow;
                        sItem.SEntryID = _result;
                    }
                    else
                        lTaskItems.Add(new cTaskItem { Touched = true, SEntryID = _result, OModifiedDate = DateTime.UtcNow, oItem = oItem });

                    clsSuiteCRMHelper.WriteLog("    date_start= " + str + ", date_due=" + str2);

                }
                catch (Exception ex)
                {
                    clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.AddTaskToS");
                }
            }
        }
        void TItems_ItemRemove()
        {
            if (IsTaskView)
            {
                if (sDelTaskId != "")
                {
                    try
                    {
                        foreach (var oItem in lTaskItems)
                        {
                            try
                            {
                                string sID = oItem.oItem.EntryID;
                            }
                            catch (COMException ex)
                            {
                                eNameValue[] data = new eNameValue[2];
                                data[0] = clsSuiteCRMHelper.SetNameValuePair("id", oItem.SEntryID);
                                data[1] = clsSuiteCRMHelper.SetNameValuePair("deleted", "1");
                                clsSuiteCRMHelper.SetEntry(data, "Tasks");
                                oItem.Delete = true;
                            }
                        }
                        lTaskItems.RemoveAll(a => a.Delete);                        
                    }
                    catch (Exception ex)
                    {
                        clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.TItems_ItemRemove");
                    }
                }
                sDelTaskId = "";
                
            }
        }
        
        private void StartCalendarSync()
        {
            try
            {
                Outlook.NameSpace oNS = this.Application.GetNamespace("mapi");
                if (oNS.Categories["SuiteCRM"] == null)
                {
                    oNS.Categories.Add("SuiteCRM", Outlook.OlCategoryColor.olCategoryColorGreen, Outlook.OlCategoryShortcutKey.olCategoryShortcutKeyNone);
                }
                Outlook.MAPIFolder appointmentsFolder = GetDefaultFolder("appointments");
                Outlook.Items items = appointmentsFolder.Items;
                
                items.ItemAdd -= Items_ItemAdd;
                items.ItemChange -= Items_ItemChange;
                items.ItemRemove -= Items_ItemRemove;
                items.ItemAdd += Items_ItemAdd;
                items.ItemChange += Items_ItemChange;
                items.ItemRemove += Items_ItemRemove;

                GetOutlookCalItems(appointmentsFolder);
                SyncMeetings(appointmentsFolder, "Meetings");
                SyncMeetings(appointmentsFolder, "Calls");                
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.StartCalendarSync");
            }
        }

        void Items_ItemRemove()
        {
            try
            {
                if (IsCalendarView)
                {
                    try
                    {
                        foreach (var oItem in lCalItems)
                        {
                            try
                            {
                                string sID = oItem.oItem.EntryID;
                            }
                            catch(COMException ex)
                            {
                                eNameValue[] data = new eNameValue[2];
                                data[0] = clsSuiteCRMHelper.SetNameValuePair("id", oItem.SEntryID);
                                data[1] = clsSuiteCRMHelper.SetNameValuePair("deleted", "1");
                                clsSuiteCRMHelper.SetEntry(data, oItem.SType);
                                oItem.Delete = true;
                            }
                        }                        
                        lCalItems.RemoveAll(a => a.Delete);
                    }
                    catch
                    { }
                }
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.Items_ItemRemove");
            }
        }

        void Items_ItemChange(object Item)
        {
            clsSuiteCRMHelper.WriteLog("Items_ItemChange");
            try
            {
                var aItem = Item as Outlook.AppointmentItem;

                string entryId = aItem.EntryID;
                cAppItem callitem = lCalItems.FirstOrDefault(a => a.oItem.EntryID == entryId);
                clsSuiteCRMHelper.WriteLog("CalItem EntryID=  " + aItem.EntryID);
                if (callitem != default(cAppItem))
                {
                    var utcNow = DateTime.UtcNow;
                    if (Math.Abs((int)(utcNow - callitem.OModifiedDate).TotalSeconds) > 5)
                    {
                        clsSuiteCRMHelper.WriteLog("2 callitem.IsUpdate = " + callitem.IsUpdate);
                        callitem.IsUpdate = 0;
                    }

                    clsSuiteCRMHelper.WriteLog("Before UtcNow - callitem.OModifiedDate= " + (int)(utcNow - callitem.OModifiedDate).TotalSeconds);

                    if (Math.Abs((int)(utcNow - callitem.OModifiedDate).TotalSeconds) > 2 && callitem.IsUpdate == 0)
                    {
                        callitem.OModifiedDate = DateTime.UtcNow;
                        clsSuiteCRMHelper.WriteLog("1 callitem.IsUpdate = " + callitem.IsUpdate);
                        callitem.IsUpdate++;
                    }

                    clsSuiteCRMHelper.WriteLog("callitem = " + callitem.oItem.Subject);
                    clsSuiteCRMHelper.WriteLog("callitem.SEntryID = " + callitem.SEntryID);
                    clsSuiteCRMHelper.WriteLog("callitem mod_date= " + callitem.OModifiedDate.ToString());
                    clsSuiteCRMHelper.WriteLog("utcNow= " + DateTime.UtcNow.ToString());
                    clsSuiteCRMHelper.WriteLog("UtcNow - callitem.OModifiedDate= " + (int)(DateTime.UtcNow - callitem.OModifiedDate).TotalSeconds);
                }
                else
                {
                    clsSuiteCRMHelper.WriteLog("not found callitem ");
                }


                if (IsCalendarView && lCalItems.Exists(a => a.oItem.EntryID == aItem.EntryID
                                 && callitem.IsUpdate == 1
                                 )
                )
                {
                    Outlook.UserProperty oProp = aItem.UserProperties["SType"];
                    Outlook.UserProperty oProp1 = aItem.UserProperties["SEntryID"];
                    if (oProp != null && oProp1 != null)
                    {
                        callitem.IsUpdate++;
                        AddAppointmentToS(aItem, oProp.Value.ToString(), oProp1.Value.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.Items_ItemChange");
            }
        }

        void Items_ItemAdd(object Item)
        {
            clsSuiteCRMHelper.WriteLog("Items_ItemAdd");
            var aItem = Item as Outlook.AppointmentItem;
            if (IsCalendarView && !lCalItems.Exists(a => a.oItem.EntryID == aItem.EntryID))
            {
                AddAppointmentToS(aItem, "Meetings");
            }
        }

        private void GetOutlookCalItems(Outlook.MAPIFolder appointmentsFolder)
        {
            try
            {
                if (lCalItems == null)
                {
                    lCalItems = new List<cAppItem>();
                    Outlook.Items items = appointmentsFolder.Items; //.Restrict("[MessageClass] = 'IPM.Appointment'" + GetStartDateString());
                    foreach (Outlook.AppointmentItem aItem in items)
                    {
                        if (aItem.Start < DateTime.Now.AddDays(-5))
                            continue;
                        Outlook.UserProperty oProp = aItem.UserProperties["SOModifiedDate"];
                        if (oProp != null)
                        {
                            Outlook.UserProperty oProp1 = aItem.UserProperties["SType"];
                            Outlook.UserProperty oProp2 = aItem.UserProperties["SEntryID"];
                            lCalItems.Add(new cAppItem
                            {
                                oItem = aItem,
                                OModifiedDate = DateTime.UtcNow,
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
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.GetOutlookCalItems");
            }
        }

        public DateTime GetStartDate()
        {
            DateTime dtRet = DateTime.Now.AddDays(-5);
            return new DateTime(dtRet.Year, dtRet.Month, dtRet.Day, 0, 0, 0);
        }
        public string GetStartDateString()
        {
            return " AND [Start] >='" + GetStartDate().ToString("MM/dd/yyyy HH:mm") + "'";
        }

        private void SetRecepients(Outlook.AppointmentItem aItem, string sMeetingID, string sModule)
        {
            aItem.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
            int iCount = aItem.Recipients.Count;
            for (int iItr = 1; iItr <= iCount; iItr++)
            {
                aItem.Recipients.Remove(1);
            }

            eEntryValue[] Users;
            string [] invitee_categories = {"users", "contacts", "leads"};
            foreach (string invitee_category in invitee_categories)
            {
                Users = clsSuiteCRMHelper.getRelationships(sModule, sMeetingID, invitee_category, new string[] { "id", "email1", "phone_work" });
                if (Users != null)
                {

                    foreach (var oResult1 in Users)
                    {
                        dynamic dResult1 = JsonConvert.DeserializeObject(oResult1.name_value_object.ToString());

                        clsSuiteCRMHelper.WriteLog("-------------------SetRecepients-----Start-----dResult1---2-------");
                        clsSuiteCRMHelper.WriteLog(Convert.ToString(dResult1));
                        clsSuiteCRMHelper.WriteLog("-------------------SetRecepients-----End---------------");

                       /* clsSuiteCRMHelper.WriteLog("-------------------SetRecepients GetAttendeeList-----Start---------------");

                        string findmeet = clsSuiteCRMHelper.getRelationship("Contacts", oResult1.id, "meetings");
                        clsSuiteCRMHelper.WriteLog("    findmeet=" + findmeet);

                        clsSuiteCRMHelper.WriteLog("-------------------SetRecepients GetAttendeeList-----End---------------");*/

                        string phone_work = dResult1.phone_work.value.ToString();
                        string sTemp = 
                            (sModule == "Meetings") || String.IsNullOrEmpty(phone_work) || String.IsNullOrWhiteSpace(phone_work) ?
                                dResult1.email1.value.ToString() :
                                dResult1.email1.value.ToString() + ":" + phone_work;
                        aItem.Recipients.Add(sTemp);
                    }
                }
            }
        }

        private void SetMeetings(eEntryValue[] el, Outlook.MAPIFolder appointmentsFolder, string sModule)
        {

            foreach (var oResult in el)
            {
                try
                {
                    dynamic dResult = JsonConvert.DeserializeObject(oResult.name_value_object.ToString());
                    DateTime date_start = DateTime.ParseExact(dResult.date_start.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);
                    date_start = date_start.Add(new DateTimeOffset(DateTime.Now).Offset);
                    if (date_start < GetStartDate())
                    {
                        continue;
                    }

                    var oItem = lCalItems.FirstOrDefault(a => a.SEntryID == dResult.id.value.ToString() && a.SType == sModule);
                    if (oItem == default(cAppItem))
                    {
                        Outlook.AppointmentItem aItem = appointmentsFolder.Items.Add(Outlook.OlItemType.olAppointmentItem);
                        aItem.Subject = dResult.name.value.ToString();
                        aItem.Body = dResult.description.value.ToString();
                        if (!string.IsNullOrWhiteSpace(dResult.date_start.value.ToString()))
                        {
                            aItem.Start = date_start;
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
                            clsSuiteCRMHelper.WriteLog("   default SetRecepients");
                            SetRecepients(aItem, dResult.id.value.ToString(), sModule);

                            //}
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
                        lCalItems.Add(new cAppItem
                        {
                            oItem = aItem,
                            OModifiedDate = DateTime.ParseExact(dResult.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null),
                            SType = sModule,
                            SEntryID = dResult.id.value.ToString(),
                            Touched = true
                        });                       
                        aItem.Save();
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
                                aItem.Start = date_start;
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
                                    clsSuiteCRMHelper.WriteLog("    SetRecepients");
                                    SetRecepients(aItem, dResult.id.value.ToString(), sModule);
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
                        clsSuiteCRMHelper.WriteLog("Not default dResult.date_modified= "+ dResult.date_modified.value.ToString());
                        oItem.OModifiedDate =DateTime.ParseExact(dResult.date_modified.value.ToString(),"yyyy-MM-dd HH:mm:ss", null);
                    }
                }
                catch (Exception ex)
                {
                    clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.SyncMeetings");
                }
            }
        }
        private void SyncMeetings(Outlook.MAPIFolder appointmentsFolder, string sModule)
        {
            clsSuiteCRMHelper.WriteLog("SyncMeetings");
            try
            {
                int iOffset = 0;
                while (true)
                {
                    eGetEntryListResult _result2 = clsSuiteCRMHelper.GetEntryList(sModule, "assigned_user_id = '" + clsSuiteCRMHelper.GetUserId() + "'",
                                    0, "date_start DESC", iOffset, false, clsSuiteCRMHelper.GetSugarFields(sModule));

                

                    if (_result2 != null)
                    {
                        if (iOffset == _result2.next_offset)
                            break;

                        SetMeetings(_result2.entry_list, appointmentsFolder, sModule);
                    }
                    if (iOffset == _result2.next_offset)
                        iOffset = 0;
                    else
                        iOffset = _result2.next_offset;
                    if (iOffset == 0)
                        break;
                }
                eEntryValue[] invited = clsSuiteCRMHelper.getRelationships("Users", clsSuiteCRMHelper.GetUserId(), sModule.ToLower(), clsSuiteCRMHelper.GetSugarFields(sModule));
                if (invited!=null)
                {

                    SetMeetings(invited, appointmentsFolder, sModule);
                }
                try
                {
                    if (sModule == "Meetings")
                    {
                        var lItemToBeDeletedO = lCalItems.Where(a => !a.Touched && !string.IsNullOrWhiteSpace(a.OModifiedDate.ToString()) && a.SType == sModule);
                        foreach (var oItem in lItemToBeDeletedO)
                        {
                            try
                            {
                                oItem.oItem.Delete();
                            }
                            catch (Exception ex)
                            {
                                clsSuiteCRMHelper.WriteLog("   Exception  oItem.oItem.Delete");
                            }


                        }
                        lCalItems.RemoveAll(a => !a.Touched && !string.IsNullOrWhiteSpace(a.OModifiedDate.ToString()) && a.SType == sModule);
                    }
                    var lItemToBeAddedToS = lCalItems.Where(a => !a.Touched && string.IsNullOrWhiteSpace(a.OModifiedDate.ToString()) && a.SType == sModule);
                    foreach (var oItem in lItemToBeAddedToS)
                    {
                        AddAppointmentToS(oItem.oItem, sModule);
                    }
                }
                catch (Exception ex)
                {
                    clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.SyncMeetings");
                }
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.SyncMeetings");
            }
        }
        private void AddAppointmentToS(Outlook.AppointmentItem aItem, string sModule, string sID = "")
        {
            clsSuiteCRMHelper.WriteLog("AddAppointmentToS");
            if (!settings.SyncCalendar)
                return;
            if (aItem != null)
            {
                try
                {
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
                    if (sID == "")
                    {
                        clsSuiteCRMHelper.WriteLog("    -- AddAppointmentToS AddAppointmentToS sID =" + sID);

                        eSetRelationshipValue info = new eSetRelationshipValue
                        {
                            module2 = "meetings",
                            module2_id = _result,
                            module1 = "Users",
                            module1_id = clsSuiteCRMHelper.GetUserId()
                        };
                        clsSuiteCRMHelper.SetRelationship(info);
                                                
                    }
                    if (aItem.Recipients!=null)
                    {
                        foreach (Outlook.Recipient objRecepient in aItem.Recipients)
                        {
                            try
                            {
                                clsSuiteCRMHelper.WriteLog("objRecepientName= " + objRecepient.Name.ToString());
                                clsSuiteCRMHelper.WriteLog("objRecepient= " + objRecepient.Address.ToString());
                            }
                            catch
                            {
                                clsSuiteCRMHelper.WriteLog("objRecepient ERROR");
                                continue;
                            }

                            string sCID = GetID(objRecepient.Address, "Contacts");
                            if (sCID != "")
                            {
                                //filter unnecessary contacts 
                                /*var contact = lContactItems.Where(a => a.SEntryID.ToString() == sCID).FirstOrDefault();
                                
                                if (contact == default(cContactItem))
                                    continue;*/

                                eSetRelationshipValue info = new eSetRelationshipValue
                                {
                                    module2 = "meetings",
                                    module2_id = _result,
                                    module1 = "Contacts",
                                    module1_id = sCID
                                };

                               /* clsSuiteCRMHelper.WriteLog("-------------------GetAttendeeList-----Start---------------");

                                string findmeet = clsSuiteCRMHelper.getRelationship("Meetings", _result, "contacts");
                                clsSuiteCRMHelper.WriteLog("    findmeet="+ findmeet);

                                clsSuiteCRMHelper.WriteLog("-------------------GetAttendeeList-----End---------------");*/

                                /*foreach (cContactItem lc in lContactItems)
                                {
                                    clsSuiteCRMHelper.WriteLog("    lc.SEntryID= " + lc.SEntryID.ToString()); 
                                }*/

                                clsSuiteCRMHelper.WriteLog("    SetRelationship 1");
                                clsSuiteCRMHelper.WriteLog("    sCID=" + sCID); 
                                clsSuiteCRMHelper.SetRelationship(info);

                                string AccountID = clsSuiteCRMHelper.getRelationship("Contacts", sCID, "accounts");

                                if (AccountID != "")
                                {
                                    info = new eSetRelationshipValue
                                    {
                                        module2 = "meetings",
                                        module2_id = _result,
                                        module1 = "Accounts",
                                        module1_id = AccountID
                                    };
                                    clsSuiteCRMHelper.SetRelationship(info);
                                }
                                continue;
                            }
                            sCID = GetID(objRecepient.Address, "Users");
                            if (sCID != "")
                            {
                                eSetRelationshipValue info = new eSetRelationshipValue
                                {
                                    module2 = "meetings",
                                    module2_id = _result,
                                    module1 = "Users",
                                    module1_id = sCID
                                };
                                clsSuiteCRMHelper.SetRelationship(info);
                                continue;
                            }
                            sCID = GetID(objRecepient.Address, "Leads");
                            if (sCID != "")
                            {
                                eSetRelationshipValue info = new eSetRelationshipValue
                                {
                                    module2 = "meetings",
                                    module2_id = _result,
                                    module1 = "Leads",
                                    module1_id = sCID
                                };
                                clsSuiteCRMHelper.WriteLog("    SetRelationship 2");
                                clsSuiteCRMHelper.SetRelationship(info);
                                continue;
                            }
                        }
                    }
                    Outlook.UserProperty oProp = aItem.UserProperties["SOModifiedDate"];
                    if (oProp == null)
                        oProp = aItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                    oProp.Value = DateTime.UtcNow;
                    Outlook.UserProperty oProp1 = aItem.UserProperties["SType"];
                    if (oProp1 == null)
                        oProp1 = aItem.UserProperties.Add("SType", Outlook.OlUserPropertyType.olText);
                    oProp1.Value = sModule;
                    Outlook.UserProperty oProp2 = aItem.UserProperties["SEntryID"];
                    if (oProp2 == null)
                        oProp2 = aItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                    oProp2.Value = _result;
                    clsSuiteCRMHelper.WriteLog("    AddAppointmentToS Save ");
                    aItem.Save();
                    string entryId = aItem.EntryID;
                    var sItem = lCalItems.FirstOrDefault(a => a.oItem.EntryID == entryId);
                    if (sItem != default(cAppItem))
                    {
                        sItem.oItem = aItem;
                        sItem.OModifiedDate = DateTime.UtcNow;
                        sItem.SEntryID = _result;
                        clsSuiteCRMHelper.WriteLog("    AddAppointmentToS Edit ");
                    }
                    else
                    {
                        lCalItems.Add(new cAppItem { SEntryID = _result, SType = sModule, OModifiedDate = DateTime.UtcNow, oItem = aItem });
                        clsSuiteCRMHelper.WriteLog("    AddAppointmentToS New ");
                    }
                }
                catch (Exception ex)
                {
                    clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.AddAppointmentToS");
                }
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
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.ThisAddIn_Shutdown");
            }
        }

        private void UnregisterEvents()
        {
            try
            {
                this.Application.ItemContextMenuDisplay -= new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(this.Application_ItemContextMenuDisplay);
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.UnregisterEvents");
            }
            try
            {
                this.btnArvive.Click -= new Office._CommandBarButtonEvents_ClickEventHandler(this.cbtnArchive_Click);
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.UnregisterEvents");
            }

            try
            {
                this.objExplorer.Application.NewMailEx -= new Outlook.ApplicationEvents_11_NewMailExEventHandler(this.Application_NewMail);
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.UnregisterEvents");
            }

            try
            {
                this.objExplorer.Application.ItemSend -= new Outlook.ApplicationEvents_11_ItemSendEventHandler(this.Application_ItemSend);
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.UnregisterEvents");
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
            try
            {
                Outlook.Selection selection = Selection;
                Outlook.MailItem item1 = (Outlook.MailItem)selection[1];
                Office.CommandBarButton objMainMenu = (Office.CommandBarButton)CommandBar.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, this.missing, this.missing, this.missing, this.missing);
                objMainMenu.Caption = "SuiteCRM Archive";
                objMainMenu.Visible = true;
                objMainMenu.Picture = RibbonImageHelper.Convert(Resources.SuiteCRM1);
                objMainMenu.Click += new Office._CommandBarButtonEvents_ClickEventHandler(this.contextMenuArchiveButton_Click);
            }
            catch (Exception)
            {

            }
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
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.Application_ItemSend");
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
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.Application_NewMail");
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
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.Authenticate");
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
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.GetMailFolders"); ;
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
                objMail.Categories = "SuiteCRM";
                objMail.Save();
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.ArchiveEmail");
            }
        }

        public string GetID(string sEmailID, string sModule)
        {
            string str5 = "(" + sModule.ToLower() + ".id in (select eabr.bean_id from email_addr_bean_rel eabr INNER JOIN email_addresses ea on eabr.email_address_id = ea.id where eabr.bean_module = '" + sModule + "' and ea.email_address LIKE '%" + clsGlobals.MySqlEscape(sEmailID) + "%'))";

            clsSuiteCRMHelper.WriteLog("-------------------GetID-----Start---------------");

            clsSuiteCRMHelper.WriteLog("    str5=" + str5);

            clsSuiteCRMHelper.WriteLog("-------------------GetID-----End---------------");

            string[] fields = new string[1];
            fields[0] = "id";
            eGetEntryListResult _result = clsSuiteCRMHelper.GetEntryList(sModule, str5, settings.SyncMaxRecords, "date_entered DESC", 0, false, fields);
            if (_result.result_count > 0)
            {
                return clsSuiteCRMHelper.GetValueByKey(_result.entry_list[0], "id");
            }
            return "";
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
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.ArchiveEmailThread");
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
                        clsSuiteCRMHelper.WriteException(ex1, "ThisAddIn.Base64Encode");
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
                clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.ArchiveFolderItems");
            }
        }

        public void ProcessMails(DateTime? dtAutoArchiveFrom = null)
        {
            if (settings.AutoArchive == false)
                return;
            System.Threading.Thread.Sleep(5000);
            while (true)
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
                    clsSuiteCRMHelper.WriteException(ex, "ThisAddIn.ProcessMails");
                }
                if (dtAutoArchiveFrom != null)
                    break;

                System.Threading.Thread.Sleep(5000);
            }
        }

        private TimeSpan[] ParseTimesFromTaskBody(string body)
        {
            try
            {
                if (string.IsNullOrEmpty(body))
                    return null;
                TimeSpan[] timesToAdd = new TimeSpan[2];
                List<int> hhmm = new List<int>(4);

                string times = body.Substring(body.LastIndexOf("#<")).Substring(2);
                char[] sep = { '<', '#', ':' };
                int parsed = 0;
                foreach (var digit in times.Split(sep))
                {
                    int.TryParse(digit, out parsed);
                    hhmm.Add(parsed);
                    parsed = 0;
                }

                TimeSpan start_time = TimeSpan.FromHours(hhmm[0]).Add(TimeSpan.FromMinutes(hhmm[1]));
                TimeSpan due_time = TimeSpan.FromHours(hhmm[2]).Add(TimeSpan.FromMinutes(hhmm[3]));
                timesToAdd[0] = start_time;
                timesToAdd[1] = due_time;
                return timesToAdd;
            }
            catch
            {
                clsSuiteCRMHelper.WriteLog("Body doesn't have time string");
                return null;
            }
           
        }
    }
}
