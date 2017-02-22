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
namespace SuiteCRMAddIn.BusinessLogic
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Newtonsoft.Json;
    using SuiteCRMClient;
    using SuiteCRMClient.RESTObjects;
    using SuiteCRMClient.Logging;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System.Runtime.InteropServices;

    public class ContactSyncing: Syncing<Outlook.ContactItem>
    {
        public ContactSyncing(SyncContext context)
            : base(context)
        {
        }

        public override bool SyncingEnabled => settings.SyncContacts;

        public void StartSync()
        {
            try
            {
                Log.Info("ContactSync thread starting");
                Outlook.NameSpace oNS = this.Application.GetNamespace("mapi");
                Outlook.MAPIFolder folder = GetDefaultFolder();

                GetOutlookItems(folder);
                SyncFolder(folder);
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.StartContactSync", ex);
            }
            finally
            {
                Log.Info("ContactSync thread completed");
            }
        }

        /// <summary>
        /// Synchronise items in the specified folder with the specified SuiteCRM module.
        /// </summary>
        /// <remarks>
        /// TODO: candidate for refactoring upwards, in concert with AppointmentSyncing.SyncFolder.
        /// </remarks>
        /// <param name="folder">The folder.</param>
        private void SyncFolder(Outlook.MAPIFolder folder)
        {
            Log.Info($"ContactSyncing.SyncFolder: '{folder}'");
            try
            {
                if (HasAccess("Contacts", "export"))
                {
                    var untouched = new HashSet<SyncState<Outlook.ContactItem>>(ItemsSyncState);
                    int nextOffset = -1; // offset of the next page of entries, if any.

                    for (int iOffset = 0; iOffset != nextOffset; iOffset = nextOffset)
                    {
                        eGetEntryListResult entriesPage = clsSuiteCRMHelper.GetEntryList("Contacts",
                                        "contacts.assigned_user_id = '" + clsSuiteCRMHelper.GetUserId() + "'",
                                        0, "date_entered DESC", iOffset, false, clsSuiteCRMHelper.GetSugarFields("Contacts"));
                        nextOffset = entriesPage.next_offset;
                        if (iOffset != nextOffset)
                        {
                            UpdateItemsFromCrmToOutlook(entriesPage.entry_list, folder, untouched);
                        }
                    }
                    try
                    {
                        // Create the lists first, because deleting items changes the value of 'ExistedInCrm'.
                        var syncableButNotOnCrm = untouched.Where(s => s.ShouldSyncWithCrm);
                        var toDeleteFromOutlook = syncableButNotOnCrm.Where(a => a.ExistedInCrm).ToList();
                        var toCreateOnCrmServer = syncableButNotOnCrm.Where(a => !a.ExistedInCrm).ToList();

                        foreach (var item in toDeleteFromOutlook)
                        {
                            item.OutlookItem.Delete();
                            LogItemAction(item.OutlookItem, "AppointmentSyncing.RemoveItemSyncState, removed item from queue");
                            ItemsSyncState.Remove(item);
                        }

                        foreach (var oItem in toCreateOnCrmServer)
                        {
                            AddToCrm(oItem.OutlookItem);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error("ContactSyncing.SyncContacts", ex);
                    }
                }
                else
                {
                    Log.Warn("ContactSyncing.SyncContacts: CRM server denied access to export Contacts");
                }
            }
            catch (Exception ex)
            {
                Log.Error("ContactSyncing.SyncContacts", ex);
            }
        }

        /// <summary>
        /// Remove an outlook item and its associated sync state.
        /// </summary>
        /// <remarks>
        /// TODO: candidate for refactoring to superclass.
        /// </remarks>
        /// <param name="syncState">The sync state of the item to remove.</param>
        private void RemoveItemAndSyncState(SyncState<Outlook.ContactItem> syncState)
        {
            this.LogItemAction(syncState.OutlookItem, "ContactSyncing.SyncFolder, deleting item");
            try
            {
                syncState.OutlookItem.Delete();
            }
            catch (Exception ex)
            {
                Log.Error("ContactSyncing.SyncFolder: Exception  oItem.oItem.Delete", ex);
            }
            this.RemoveItemSyncState(syncState);
        }

        /// <summary>
        /// Remove an item from ItemsSyncState.
        /// </summary>
        /// <remarks>
        /// TODO: candidate for refactoring to superclass.
        /// </remarks>
        /// <param name="item">The sync state of the item to remove.</param>
        private void RemoveItemSyncState(SyncState<Outlook.ContactItem> item)
        {
            this.LogItemAction(item.OutlookItem, "AppointmentSyncing.RemoveItemSyncState, removed item from queue");
            this.ItemsSyncState.Remove(item);
        }

        /// <summary>
        /// Update these items.
        /// TODO: This is a candidate for refactoring with AppointmentSyncing.UpdateItemsFromCrmToOutlook
        /// </summary>
        /// <param name="items">The items to be synchronised.</param>
        /// <param name="folder">The outlook folder to synchronise into.</param>
        /// <param name="untouched">A list of items which have not yet been synchronised; this list is 
        /// modified (destructuvely changed) by the action of this method.</param>
        private void UpdateItemsFromCrmToOutlook(
            eEntryValue[] items, 
            Outlook.MAPIFolder folder, 
            HashSet<SyncState<Outlook.ContactItem>> untouched)
        {
            foreach (var oResult in items)
            {
                try
                {
                    var state = UpdateFromCrm(folder, oResult);
                    if (state != null)
                    {
                        untouched.Remove(state);
                        LogItemAction(state.OutlookItem, "ContactSyncing.UpdateAppointmentsFromCrmToOutlook, item removed from untouched");
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("ContactSyncing.UpdateItemsFromCrmToOutlook", ex);
                }
            }
        }

        private SyncState<Outlook.ContactItem> UpdateFromCrm(Outlook.MAPIFolder folder, eEntryValue candidateItem)
        {
            SyncState<Outlook.ContactItem> result;
            dynamic crmItem = JsonConvert.DeserializeObject(candidateItem.name_value_object.ToString());
            String id = crmItem.id.value.ToString();
            var oItem = ItemsSyncState.FirstOrDefault(a => a.CrmEntryId == crmItem.id.value.ToString());

            if (ShouldSyncContact(crmItem))
            {
                Log.Info(
                    string.Format(
                        "ContactSyncing.UpdateFromCrm, entry id is '{0}', sync_contact is true, syncing",
                        id));

                if (oItem == null)
                {
                    result = AddNewItemFromCrmToOutlook(folder, crmItem);
                }
                else
                {
                    result = UpdateExistingOutlookItemFromCrm(crmItem, oItem);
                }
            }
            else
            {
                Log.Info(
                    string.Format(
                        "ContactSyncing.UpdateFromCrm, entry id is '{0}', sync_contact is false, not syncing",
                        id));
                
                result = oItem;
            }

            return result;
        }

        /// <summary>
        /// Add an item existing in CRM but not found in Outlook to Outlook.
        /// </summary>
        /// <param name="appointmentsFolder">The Outlook folder in which the item should be stored.</param>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <returns>A sync state object for the new item.</returns>
        private SyncState<Outlook.ContactItem> AddNewItemFromCrmToOutlook(Outlook.MAPIFolder contactFolder, dynamic crmItem)
        {
            Log.Info(
                (string)string.Format(
                    "ContactSyncing.AddNewItemFromCrmToOutlook, entry id is '{0}', creating in Outlook.",
                    crmItem.id.value.ToString()));

            Outlook.ContactItem olItem = ConstructOutlookItemFromCrmItem(contactFolder, crmItem);
            var newState = new ContactSyncState
            {
                OutlookItem = olItem,
                OModifiedDate = DateTime.ParseExact(crmItem.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null),
                CrmEntryId = crmItem.id.value.ToString(),
            };
            ItemsSyncState.Add(newState);
            olItem.Save();

            LogItemAction(newState.OutlookItem, "AppointmentSyncing.AddNewItemFromCrmToOutlook, saved item");

            return newState;
        }

        /// <summary>
        /// Log a message regarding this Outlook appointment.
        /// </summary>
        /// <param name="olItem">The outlook item.</param>
        /// <param name="message">The message to be logged.</param>
        private void LogItemAction(Outlook.ContactItem olItem, string message)
        {
            try
            {
                Outlook.UserProperty olPropertyEntryId = olItem.UserProperties["SEntryID"];
                string crmId = olPropertyEntryId == null ?
                    "[not present]" :
                    olPropertyEntryId.Value;
                Log.Info(
                    String.Format("{0}:\n\tOutlook Id  : {1}\n\tCRM Id      : {2}\n\tFull name   : '{3}'\n\tSensitivity : {4}",
                    message, olItem.EntryID, crmId, olItem.FullName, olItem.Sensitivity));
            }
            catch (COMException)
            {
                // Ignore: happens if the outlook item is already deleted.
            }
        }

        /// <summary>
        /// Return true if this CRM contact should be synchronised with Outlook.
        /// </summary>
        /// <param name="crmContact">The CRM contact.</param>
        /// <returns>true if this CRM contact should be synchronised with Outlook.</returns>
        private bool ShouldSyncContact(dynamic crmContact)
        {
            bool result = false;
            String stringValue = crmContact.sync_contact.value.ToString();

            try
            {
                result = Boolean.Parse(stringValue);
            }
            catch (FormatException fex)
            {
                Log.Warn(
                    String.Format(
                        "ContactSyncing.ShouldSyncContact: unexpected format for sync_contact, '{0}'", 
                        stringValue));
            }

            return result;
        }

        /// <summary>
        /// Update an existing Outlook item with values taken from a corresponding CRM item. Note that 
        /// this just overwrites all values in the Outlook item.
        /// </summary>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <param name="oItem">The outlook item assumed to correspond with the CRM item.</param>
        /// <returns>An appropriate sync state.</returns>
        private SyncState<Outlook.ContactItem> UpdateExistingOutlookItemFromCrm(dynamic crmItem, SyncState<Outlook.ContactItem> oItem)
        {
            Outlook.ContactItem cItem = oItem.OutlookItem;
            Outlook.UserProperty oProp = cItem.UserProperties["SOModifiedDate"];

            if (oProp.Value != crmItem.date_modified.value.ToString())
            {
                cItem.FirstName = crmItem.first_name.value.ToString();
                cItem.LastName = crmItem.last_name.value.ToString();
                cItem.Email1Address = crmItem.email1.value.ToString();
                cItem.BusinessTelephoneNumber = crmItem.phone_work.value.ToString();
                cItem.HomeTelephoneNumber = crmItem.phone_home.value.ToString();
                cItem.MobileTelephoneNumber = crmItem.phone_mobile.value.ToString();
                cItem.JobTitle = crmItem.title.value.ToString();
                cItem.Department = crmItem.department.value.ToString();
                cItem.BusinessAddressCity = crmItem.primary_address_city.value.ToString();
                cItem.BusinessAddressCountry = crmItem.primary_address_country.value.ToString();
                cItem.BusinessAddressPostalCode = crmItem.primary_address_postalcode.value.ToString();

                if (crmItem.primary_address_street.value != null)
                    cItem.BusinessAddressStreet = crmItem.primary_address_street.value.ToString();
                cItem.Body = crmItem.description.value.ToString();
                cItem.Account = cItem.CompanyName = "";
                if (crmItem.account_name != null && crmItem.account_name.value != null)
                {
                    cItem.Account = crmItem.account_name.value.ToString();
                    cItem.CompanyName = crmItem.account_name.value.ToString();
                }

                cItem.BusinessFaxNumber = crmItem.phone_fax.value.ToString();
                cItem.Title = crmItem.salutation.value.ToString();
                if (oProp == null)
                    oProp = cItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);
                oProp.Value = crmItem.date_modified.value.ToString();
                Outlook.UserProperty oProp2 = cItem.UserProperties["SEntryID"];
                if (oProp2 == null)
                    oProp2 = cItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                oProp2.Value = crmItem.id.value.ToString();
                Log.Info("    save not default");
                Log.Info(cItem.FullName + "     is saving with" + cItem.Sensitivity);
                cItem.Save();

                LogItemAction(cItem, "AppointmentSyncing.UpdateExistingOutlookItemFromCrm, saved item");
            }
            Log.Info($"AppointmentSyncing.UpdateExistingOutlookItemFromCrm: {cItem.FullName} dResult.date_modified={crmItem.date_modified}");
            oItem.OModifiedDate = DateTime.ParseExact(crmItem.date_modified.value.ToString(), "yyyy-MM-dd HH:mm:ss", null);
            return oItem;
        }

        private Outlook.ContactItem ConstructOutlookItemFromCrmItem(Outlook.MAPIFolder contactFolder, dynamic dResult)
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

            LogItemAction(cItem, "AppointmentSyncing.ConstructOutlookItemFromCrmItem");
            return cItem;
        }

        private void GetOutlookItems(Outlook.MAPIFolder taskFolder)
        {
            try
            {
                if (ItemsSyncState == null)
                {
                    ItemsSyncState = new List<SyncState<Outlook.ContactItem>>();
                    Outlook.Items items = taskFolder.Items.Restrict("[MessageClass] = 'IPM.Contact'");
                    foreach (Outlook.ContactItem oItem in items)
                    {
                        AddOrGetSyncState(oItem);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.GetOutlookCItems", ex);
            }
        }

        override protected void OutlookItemChanged(Outlook.ContactItem item)
        {
            if (item != null) SaveChangedItem(item);
        }

        private void SaveChangedItem(Outlook.ContactItem oItem)
        {
            var contact = AddOrGetSyncState(oItem);
            if (!ShouldPerformSyncNow(contact)) return;
            if (contact.ShouldSyncWithCrm)
            {
                if (contact.ExistedInCrm)
                {
                    contact.IsUpdate = 2;
                    AddToCrm(oItem, contact.CrmEntryId);
                }
                else
                {
                    AddToCrm(oItem);
                }
            }
            else
            {
                RemoveFromCrm(contact);
            }
        }

        /// <summary>
        /// TODO: I (AF) do not understand the purpose of this logic. (Pre-existing code, slightly cleaned-up.)
        /// </summary>
        /// <param name="contact"></param>
        /// <returns></returns>
        private bool ShouldPerformSyncNow(SyncState<Outlook.ContactItem> contact)
        {
            var modifiedSinceSeconds = Math.Abs((DateTime.UtcNow - contact.OModifiedDate).TotalSeconds);
            if (modifiedSinceSeconds > 5 || modifiedSinceSeconds > 2 && contact.IsUpdate == 0)
            {
                contact.OModifiedDate = DateTime.UtcNow;
                contact.IsUpdate = 1;
            }

            return (IsCurrentView && contact.IsUpdate == 1);
        }

        override protected void OutlookItemAdded(Outlook.ContactItem item)
        {
            if (IsCurrentView && item != null)
                AddNewItem(item);
        }

        private void AddNewItem(Outlook.ContactItem item)
        {
            var state = AddOrGetSyncState(item);
            if (state.ShouldSyncWithCrm)
            {
                AddToCrm(item, state.CrmEntryId);
            }
            else
            {
                Log.Info($"Ignoring addition of {item.FullName} because it is {item.Sensitivity}");
            }
        }

        private void AddToCrm(Outlook.ContactItem oItem, string sID = null)
        {
            if (!SyncingEnabled)
                return;
            if (oItem == null) return;
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

                if (string.IsNullOrEmpty(sID))
                    data[17] = clsSuiteCRMHelper.SetNameValuePair("assigned_user_id", clsSuiteCRMHelper.GetUserId());
                else
                    data[17] = clsSuiteCRMHelper.SetNameValuePair("id", sID);

                _result = clsSuiteCRMHelper.SetEntryUnsafe(data, "Contacts");
                Outlook.UserProperty oProp = oItem.UserProperties["SOModifiedDate"];
                if (oProp == null)
                    oProp = oItem.UserProperties.Add("SOModifiedDate", Outlook.OlUserPropertyType.olText);

                oProp.Value = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss");

                Outlook.UserProperty oProp2 = oItem.UserProperties["SEntryID"];
                if (oProp2 == null)
                    oProp2 = oItem.UserProperties.Add("SEntryID", Outlook.OlUserPropertyType.olText);
                oProp2.Value = _result;

                Log.Info(oItem.FullName + " from save Sensitivity= " + oItem.Sensitivity);

                Log.Info("        Save");
                oItem.Save();

                var state = AddOrGetSyncState(oItem);

                state.OModifiedDate = DateTime.UtcNow;
                state.CrmEntryId = _result;
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.AddContactToS", ex);
            }
        }

        private SyncState<Outlook.ContactItem> AddOrGetSyncState(Outlook.ContactItem oItem)
        {
            var entryId = oItem.EntryID;
            var existingState = ItemsSyncState.FirstOrDefault(a => a.OutlookItem.EntryID == entryId);
            if (existingState != null)
            {
                existingState.OutlookItem = oItem;
                return existingState;
            }
            else
            {
                var newState = new ContactSyncState
                {
                    OutlookItem = oItem,
                    CrmEntryId = oItem.UserProperties["SEntryID"]?.Value.ToString(),
                    OModifiedDate = ParseDateTimeFromUserProperty(oItem.UserProperties["SOModifiedDate"]?.Value.ToString()),
                };
                ItemsSyncState.Add(newState);
                return newState;
            }
        }

        public override Outlook.MAPIFolder GetDefaultFolder()
        {
            return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
        }

        protected override bool IsCurrentView => Context.CurrentFolderItemType == Outlook.OlItemType.olContactItem;

        // Should presumably be removed at some point. Existing code was ignoring deletions for Contacts and Tasks
        // (but not for Appointments).
        protected override bool PropagatesLocalDeletions => true;
    }
}
