﻿/**
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
namespace SuiteCRMAddIn.BusinessLogic
{
    using ProtoItems;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using Outlook = Microsoft.Office.Interop.Outlook;

    public class ContactSyncing: Synchroniser<Outlook.ContactItem>
    {
        /// <summary>
        /// The module I synchronise with.
        /// </summary>
        public const string CrmModule = "Contacts";

        public ContactSyncing(string name, SyncContext context)
            : base(name, context)
        {
            this.fetchQueryPrefix = "contacts.assigned_user_id = '{0}'";
        }

        public override SyncDirection.Direction Direction => settings.SyncContacts;

        /// <summary>
        /// The actual transmission lock object of this synchroniser.
        /// </summary>
        private object txLock = new object();

        /// <summary>
        /// Allow my parent class to access my transmission lock object.
        /// </summary>
        protected override object TransmissionLock
        {
            get
            {
                return txLock;
            }
        }

        public override string DefaultCrmModule
        {
            get
            {
                return ContactSyncing.CrmModule;
            }
        }

        protected override void SaveItem(Outlook.ContactItem olItem)
        {
            olItem.Save();
        }

        /// <summary>
        /// Synchronise items in the specified folder with the specified SuiteCRM module.
        /// </summary>
        /// <param name="folder">The folder.</param>
        /// <param name="crmModule">The module to snychronise with.</param>
        protected override void SyncFolder(Outlook.MAPIFolder folder, string crmModule)
        {
            Log.Info($"ContactSyncing.SyncFolder: '{folder}'");
            try
            {
                if (this.permissionsCache.HasExportAccess())
                {
                    var untouched = new HashSet<SyncState<Outlook.ContactItem>>(ItemsSyncState);

                    MergeRecordsFromCrm(folder, crmModule, untouched);

                    try
                    {
                        var syncableButNotOnCrm = untouched.Where(s => s.ShouldSyncWithCrm);
                        ResolveUnmatchedItems(syncableButNotOnCrm);
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

        protected override SyncState<Outlook.ContactItem> AddOrUpdateItemFromCrmToOutlook(Outlook.MAPIFolder folder, string crmType, eEntryValue crmItem)
        {
            SyncState<Outlook.ContactItem> result;

            String id = crmItem.GetValueAsString("id");
            SyncState<Outlook.ContactItem> syncStateForItem = GetExistingSyncState(crmItem);

            if (ShouldSyncContact(crmItem))
            {
                Log.Info(
                    string.Format(
                        "ContactSyncing.UpdateFromCrm, entry id is '{0}', sync_contact is true, syncing",
                        id));

                if (syncStateForItem == null)
                {
                    result = AddNewItemFromCrmToOutlook(folder, crmItem);
                }
                else
                {
                    result = UpdateExistingOutlookItemFromCrm(crmItem, syncStateForItem);
                }
            }
            else if (syncStateForItem != null &&
                syncStateForItem.OutlookItem != null)
            {
                /* The date_modified value in CRM does not get updated when the sync_contact value
                 * is changed. But seeing this value can only be updated at the CRM side, if it
                 * has changed the change must have been at the CRM side. It doesn't change to false, 
                 * it simply ceases to be sent. Set the item to Private in Outlook. */
                if (syncStateForItem.OutlookItem.Sensitivity != Outlook.OlSensitivity.olPrivate)
                {
                    Log.Info($"ContactSyncing.UpdateFromCrm: setting sensitivity of contact {crmItem.GetValueAsString("first_name")} {crmItem.GetValueAsString("last_name")} ({crmItem.GetValueAsString("email1")}) to private");
                    syncStateForItem.OutlookItem.Sensitivity = Outlook.OlSensitivity.olPrivate;
                }

                result = syncStateForItem;
            }
            else
            {
                Log.Info(
                    string.Format(
                        "ContactSyncing.UpdateFromCrm, entry id is '{0}', sync_contact is false, not syncing",
                        id));

                result = syncStateForItem;
            }

            return result;
        }

        /// <summary>
        /// Detect whether the should sync flag value is different between these two representations.
        /// </summary>
        /// <param name="outlookItem">An outlook item.</param>
        /// <param name="crmItem">A CRM item, presumed to represent the same entity.</param>
        /// <returns>True if the should sync flag values are different, else false.</returns>
        private bool ShouldSyncFlagChanged(Outlook.ContactItem outlookItem, eEntryValue crmItem)
        {
            bool result = false;
            Outlook.UserProperty shouldSyncProp = outlookItem.UserProperties["SShouldSync"];

            if (shouldSyncProp != null)
            {
                string crmShouldSync = ShouldSyncContact(crmItem).ToString().ToLower();
                string olShouldSync = shouldSyncProp.Value.ToLower();

                result = crmShouldSync != olShouldSync;
            }

            return result;
        }

        /// <summary>
        /// Add an item existing in CRM but not found in Outlook to Outlook.
        /// </summary>
        /// <param name="appointmentsFolder">The Outlook folder in which the item should be stored.</param>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <returns>A sync state object for the new item.</returns>
        private SyncState<Outlook.ContactItem> AddNewItemFromCrmToOutlook(Outlook.MAPIFolder contactFolder, eEntryValue crmItem)
        {
            Log.Info(
                (string)string.Format(
                    "ContactSyncing.AddNewItemFromCrmToOutlook, entry id is '{0}', creating in Outlook.",
                    crmItem.GetValueAsString("id")));

            Outlook.ContactItem olItem = contactFolder.Items.Add(Outlook.OlItemType.olContactItem);

            this.SetOutlookItemPropertiesFromCrmItem(crmItem, olItem);

            var newState = new ContactSyncState
            {
                OutlookItem = olItem,
                OModifiedDate = DateTime.ParseExact(crmItem.GetValueAsString("date_modified"), "yyyy-MM-dd HH:mm:ss", null),
                CrmEntryId = crmItem.GetValueAsString("id"),
            };
            ItemsSyncState.Add(newState);
            olItem.Save();

            LogItemAction(newState.OutlookItem, "AppointmentSyncing.AddNewItemFromCrmToOutlook, saved item");

            return newState;
        }

        /// <summary>
        /// Log a message regarding this Outlook item, with detail of the item.
        /// </summary>
        /// <param name="olItem">The outlook item.</param>
        /// <param name="message">The message to be logged.</param>
        protected override void LogItemAction(Outlook.ContactItem olItem, string message)
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
        /// <remarks>
        /// If the 'Sync to Outlook' field is set in CRM, we get 'true' as the value of crmItem.sync_contact. 
        /// But if the field is not set, we do not (or do not reliably) get 'false'. The sync_contact 
        /// property may have a value of ''.
        /// </remarks>
        /// <param name="crmContact">The CRM contact.</param>
        /// <returns>true if this CRM contact should be synchronised with Outlook.</returns>
        private bool ShouldSyncContact(eEntryValue crmContact)
        {
            object val = crmContact.GetValue("sync_contact");
            return Boolean.TrueString.ToLower().Equals(val.ToString().ToLower());
        }

        /// <summary>
        /// A CRM item is perceived to have changed if its modified date is different from
        /// that of its Outlook representation, or if its should sync flag is.
        /// </summary>
        /// <param name="crmItem">A CRM item.</param>
        /// <param name="outlookItem">An Outlook item, assumed to represent the same entity.</param>
        /// <returns>True if either of these propertyies differ between the representations.</returns>
        private bool CrmItemChanged(eEntryValue crmItem, Outlook.ContactItem outlookItem)
        {
            Outlook.UserProperty dateModifiedProp = outlookItem.UserProperties["SOModifiedDate"];

            return (dateModifiedProp.Value != crmItem.GetValueAsString("date_modified") ||
                ShouldSyncFlagChanged(outlookItem, crmItem));
        }

        /// <summary>
        /// Update an existing Outlook item with values taken from a corresponding CRM item. Note that 
        /// this just overwrites all values in the Outlook item.
        /// </summary>
        /// <param name="crmItem">The CRM item from which values are to be taken.</param>
        /// <param name="itemSyncState">The sync state of an outlook item assumed to correspond with the CRM item.</param>
        /// <returns>An appropriate sync state.</returns>
        private SyncState<Outlook.ContactItem> UpdateExistingOutlookItemFromCrm(eEntryValue crmItem, SyncState<Outlook.ContactItem> itemSyncState)
        {
            if (!itemSyncState.IsDeletedInOutlook)
            {
                Outlook.ContactItem outlookItem = itemSyncState.OutlookItem;
                Outlook.UserProperty dateModifiedProp = outlookItem.UserProperties["SOModifiedDate"];
                Outlook.UserProperty shouldSyncProp = outlookItem.UserProperties["SShouldSync"];
                this.LogItemAction(outlookItem, "ContactSyncing.UpdateExistingOutlookItemFromCrm");

                if (CrmItemChanged(crmItem, outlookItem))
                {
                    DateTime crmDate = DateTime.Parse(crmItem.GetValueAsString("date_modified"));
                    DateTime outlookDate = dateModifiedProp == null ? DateTime.MinValue : DateTime.Parse(dateModifiedProp.Value.ToString());

                    if (crmDate > this.LastRunCompleted && outlookDate > this.LastRunCompleted)
                    {
                        MessageBox.Show(
                            $"Contact {outlookItem.FirstName} {outlookItem.LastName} has changed both in Outlook and CRM; please check which is correct",
                            "Update problem", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else if (crmDate > outlookDate)
                    {
                        this.SetOutlookItemPropertiesFromCrmItem(crmItem, outlookItem);
                    }

                    this.LogItemAction(outlookItem, $"ContactSyncing.UpdateExistingOutlookItemFromCrm, saving with {outlookItem.Sensitivity}");

                    outlookItem.Save();
                }

                this.LogItemAction(outlookItem, "ContactSyncing.UpdateExistingOutlookItemFromCrm");
                itemSyncState.OModifiedDate = DateTime.ParseExact(crmItem.GetValueAsString("date_modified"), "yyyy-MM-dd HH:mm:ss", null);
            }
	        return itemSyncState;
        }

        /// <summary>
        /// Set all those properties of this outlook item whose values are different from the 
        /// equivalent values on this CRM item. Update the synchronisation properties only if some
        /// other property has actually changed.
        /// </summary>
        /// <param name="crmItem">The CRM item from which to take values.</param>
        /// <param name="outlookItem">The Outlook item into which to insert values.</param>
        /// <returns>true if anything was changed.</returns>
        private void SetOutlookItemPropertiesFromCrmItem(eEntryValue crmItem, Outlook.ContactItem outlookItem)
        {
            outlookItem.FirstName = crmItem.GetValueAsString("first_name");
            outlookItem.LastName = crmItem.GetValueAsString("last_name");
            outlookItem.Email1Address = crmItem.GetValueAsString("email1");
            outlookItem.BusinessTelephoneNumber = crmItem.GetValueAsString("phone_work");
            outlookItem.HomeTelephoneNumber = crmItem.GetValueAsString("phone_home");
            outlookItem.MobileTelephoneNumber = crmItem.GetValueAsString("phone_mobile");
            outlookItem.JobTitle = crmItem.GetValueAsString("title");
            outlookItem.Department = crmItem.GetValueAsString("department");
            outlookItem.BusinessAddressCity = crmItem.GetValueAsString("primary_address_city");
            outlookItem.BusinessAddressCountry = crmItem.GetValueAsString("primary_address_country");
            outlookItem.BusinessAddressPostalCode = crmItem.GetValueAsString("primary_address_postalcode");
            outlookItem.BusinessAddressState = crmItem.GetValueAsString("primary_address_state");
            outlookItem.BusinessAddressStreet = crmItem.GetValueAsString("primary_address_street");
            outlookItem.Body = crmItem.GetValueAsString("description");
            if (crmItem.GetValue("account_name") != null)
            {
                outlookItem.Account = crmItem.GetValueAsString("account_name");
                outlookItem.CompanyName = crmItem.GetValueAsString("account_name");
            }
            outlookItem.BusinessFaxNumber = crmItem.GetValueAsString("phone_fax");
            outlookItem.Title = crmItem.GetValueAsString("salutation");

            if (outlookItem.Sensitivity != Outlook.OlSensitivity.olNormal)
            {
                Log.Info($"ContactSyncing.UpdateFromCrm: setting sensitivity of contact {crmItem.GetValueAsString("first_name")} {crmItem.GetValueAsString("last_name")} ({crmItem.GetValueAsString("email1")}) to normal");
                outlookItem.Sensitivity = Outlook.OlSensitivity.olNormal;
            }

            EnsureSynchronisationPropertiesForOutlookItem(
                outlookItem, 
                crmItem.GetValueAsString("date_modified"), 
                crmItem.GetValueAsString("sync_contact"), 
                crmItem.GetValueAsString("id"));
        }

        /// <summary>
        /// Ensure that this Outlook item has a property of this name with this value.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        protected override void EnsureSynchronisationPropertyForOutlookItem(Outlook.ContactItem olItem, string name, string value)
        {
            Outlook.UserProperty olProperty = olItem.UserProperties[name];
            if (olProperty == null)
            {
                olProperty = olItem.UserProperties.Add(name, Outlook.OlUserPropertyType.olText);
            }
            olProperty.Value = value ?? string.Empty;
        }


        protected override void GetOutlookItems(Outlook.MAPIFolder taskFolder)
        {
            try
            {
                Outlook.Items items = taskFolder.Items.Restrict("[MessageClass] = 'IPM.Contact'");
                foreach (Outlook.ContactItem oItem in items)
                {
                    AddOrGetSyncState(oItem);
                }
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.GetOutlookCItems", ex);
            }
        }


        /// <summary>
        /// (Don't actually) remove the item implied by this sync state from CRM.
        /// </summary>
        /// <remarks>
        /// After considerable thought we've decided that contacts should never actually be deleted from CRM
        /// by the action of the plugin.
        /// </remarks>
        /// <param name="state">A sync state wrapping an item which has been deleted or marked private in Outlook.</param>
        protected override void RemoveFromCrm(SyncState state)
        {
            if (state is ContactSyncState)
            {
                /* which it most definitely should be */
                if (state.ExistedInCrm && (state.IsDeletedInOutlook || ! state.IsPublic))
                {
                    /* remove sync_contact relationship in CRM */
                    EnsureSyncWithOutlookSetInCRM(state.CrmEntryId, null, false);
                }
            }
            else
            {
                base.RemoveFromCrm(state);
            }
        }

        /// <summary>
        /// Add the Outlook item referenced by this sync state, which may not exist in CRM, to CRM.
        /// </summary>
        /// <param name="syncState">The sync state referencing the outlook item to add.</param>
        /// <param name="crmType">The CRM type ('module') to which it should be added</param>
        /// <param name="entryId">The id of this item in CRM, if known (in which case I should be doing
        /// an update, not an add).</param>
        /// <returns>The id of the entry added o</returns>
        internal override string AddOrUpdateItemFromOutlookToCrm(SyncState<Outlook.ContactItem> syncState, string crmType, string entryId = null)
        {
            string result = entryId;
            var outlookItem = syncState.OutlookItem;

            if (this.ShouldAddOrUpdateItemFromOutlookToCrm(outlookItem))
            {
                result = base.AddOrUpdateItemFromOutlookToCrm(syncState, crmType, entryId);

                Outlook.UserProperty syncProperty = outlookItem.UserProperties["SShouldSync"];
                string shouldSync = syncProperty == null ?
                    Boolean.TrueString.ToLower() :
                    syncProperty.Value;

                EnsureSyncWithOutlookSetInCRM(result, syncProperty);
            }
 
            return result;
        }


        /// <summary>
        /// Construct a JSON packet representing this Outlook item, and despatch it to CRM. 
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmType">The type within CRM to which the item should be added.</param>
        /// <param name="entryId">The corresponding entry id in CRM, if known.</param>
        /// <returns>The CRM id of the object created or modified.</returns>
        protected override string ConstructAndDespatchCrmItem(Outlook.ContactItem olItem, string crmType, string entryId)
        {
            return clsSuiteCRMHelper.SetEntryUnsafe(new ProtoContact(olItem).AsNameValues(entryId), crmType);
        }

        protected override SyncState<Outlook.ContactItem> ConstructSyncState(Outlook.ContactItem oItem)
        {
            return new ContactSyncState
            {
                OutlookItem = oItem,
                CrmEntryId = oItem.UserProperties["SEntryID"]?.Value.ToString(),
                OModifiedDate = ParseDateTimeFromUserProperty(oItem.UserProperties["SOModifiedDate"]?.Value.ToString()),
            };
        }


        /// <summary>
        /// If it was created in Outlook and doesn't exist in CRM,  (in which case it won't yet have a 
        /// magic SShouldSync property) then we need to guarantee changes made in CRM are copied back
        /// by setting the Sync to Outlook checkbox in CRM.
        /// </summary>
        /// <param name="contactIdInCRM">The identifier of the contact in the CRM system</param>
        /// <param name="syncProperty">If null, set the checkbox.</param>
        /// <param name="create">If provided and false, then remove rather than creating the relationship.</param>
        private static void EnsureSyncWithOutlookSetInCRM(string contactIdInCRM, Outlook.UserProperty syncProperty, bool create = true)
        {
            if (syncProperty == null)
            {
                eSetRelationshipValue info = new eSetRelationshipValue
                {
                    module1 = CrmModule,
                    module1_id = contactIdInCRM,
                    module2 = "user_sync",
                    module2_id = clsSuiteCRMHelper.GetUserId(),
                    delete = create ? 0 : 1
                };
                clsSuiteCRMHelper.SetRelationshipUnsafe(info);
            }
        }

        public override Outlook.MAPIFolder GetDefaultFolder()
        {
            return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
        }

        internal override string GetOutlookEntryId(Outlook.ContactItem olItem)
        {
            return olItem.EntryID;
        }

        internal override Outlook.OlSensitivity GetSensitivity(Outlook.ContactItem item)
        {
            return item.Sensitivity;
        }

        /// <summary>
        /// True if the currently open tab in Outlook displays items of my item type.
        /// </summary>
        /// <remarks>
        /// This is used in determining whether an item is in fact newly created by the user;
        /// it has a certain code smell to it.
        /// </remarks>
        protected override bool IsCurrentView => Context.CurrentFolderItemType == Outlook.OlItemType.olContactItem;

        /// <summary>
        /// Return the sensitivity of this outlook item.
        /// </summary>
        /// <remarks>
        /// Outlook item classes do not inherit from a common base class, so generic client code cannot refer to 'OutlookItem.Sensitivity'.
        /// </remarks>
        /// <param name="item">The outlook item whose sensitivity is required.</param>
        /// <returns>the sensitivity of the item.</returns>
        protected override bool PropagatesLocalDeletions => true;
    }
}
