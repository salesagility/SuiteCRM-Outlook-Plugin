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


namespace SuiteCRMAddIn.BusinessLogic
{
    using Exceptions;
    using Extensions;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Security.Cryptography;
    using System.Text;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// One of the problems with the design of the add-in is that we're trying to shim two 
    /// types of CRM entities (Calls, Meetings) onto one type of Outlook entity (Appointments).
    /// We need to treat them separately, and for that reason we need different sync state 
    /// classes for Calls and Meetings.
    /// </summary>
    public class SyncStateManager
    {
        /// <summary>
        /// The name of the modified date synchronisation property, which 
        /// should be updated to the date/time the item was most recently modified in Outlook or in CRM.
        /// </summary>
        public const string ModifiedDatePropertyName = "SOModifiedDate";

        /// <summary>
        /// The name of the type synchronisation property.
        /// </summary>
        public const string TypePropertyName = "SType";

        /// <summary>
        /// The name of the CRM ID synchronisation property.
        /// </summary>
        /// <see cref="SuiteCRMAddIn.Extensions.MailItemExtensions.CrmIdPropertyName"/> 
        public static string CrmIdPropertyName => string.IsNullOrWhiteSpace(Properties.Settings.Default.CurrentCrmIdPropertyName)? ConstructAndSetCrmEntryIdPropertyName() : Properties.Settings.Default.CurrentCrmIdPropertyName;

        /// <summary>
        /// Construct a new name for the CRM id property based on the hist url,
        /// set it in settings, and return it.
        /// </summary>
        /// <remarks>
        /// <para>#6661: when we change the CRM URL, we also need to change the name 
        /// of the property on which the CRM Id is held, in order neither to 
        /// use the wrong CRM ids, nor to spend a lot of time clearing old ones.</para>
        /// <para>This is a private method, which cannot be called directly from 
        /// outside this class. To force a reset of the property name, set the setting
        /// value to null or <see cref="string.Empty"/>.</para>
        /// </remarks>
        /// <seealso cref="Properties.Settings.Default.CurrentCrmIdPropertyName"/> 
        /// <returns>The value set.</returns>
        private static string ConstructAndSetCrmEntryIdPropertyName()
        {
            string previous = string.IsNullOrWhiteSpace(Properties.Settings.Default.CurrentCrmIdPropertyName) ? 
                LegacyCrmIdPropertyName : 
                Properties.Settings.Default.CurrentCrmIdPropertyName;
            byte[] bytes = MD5.Create().ComputeHash(Encoding.UTF8.GetBytes(Properties.Settings.Default.Host));

            string result = $"CrmId{BitConverter.ToString(bytes)}".Replace("-", string.Empty);
            Properties.Settings.Default.CurrentCrmIdPropertyName = result;
            Properties.Settings.Default.Save();

            Globals.ThisAddIn.Log.Info($"Updated CRM Id property name from {previous} to {result}");

            return result;
        }

        public const string LegacyCrmIdPropertyName = "SEntryID";

        /// <summary>
        /// If set, don't sync with CRM.
        /// </summary>
        public const string CRMShouldNotSyncPropertyName = "ShouldNotSyncWithCRM";

        /// <summary>
        /// My underlying instance.
        /// </summary>
        private static readonly Lazy<SyncStateManager> lazy =
            new Lazy<SyncStateManager>(() => new SyncStateManager());


        /// <summary>
        /// A lock on creating new items.
        /// </summary>
        private object creationLock = new object();

        /// <summary>
        /// A log, to log stuff to.
        /// </summary>
        private ILogger log = Globals.ThisAddIn.Log;

        /// <summary>
        /// A public accessor for my instance.
        /// </summary>
        public static SyncStateManager Instance { get { return lazy.Value; } }

        /// <summary>
        /// A dictionary of all known sync states indexed by outlook id.
        /// </summary>
        private ConcurrentDictionary<string, SyncState> byOutlookId = new ConcurrentDictionary<string, SyncState>();

        /// <summary>
        /// A dictionary of all known sync states which have global ids, indexed by global id.
        /// </summary>
        private ConcurrentDictionary<string, SyncState> byGlobalId = new ConcurrentDictionary<string, SyncState>();

        /// <summary>
        /// A dictionary of sync states indexed by crm id, where known.
        /// </summary>
        private ConcurrentDictionary<CrmId, SyncState> byCrmId = new ConcurrentDictionary<CrmId, SyncState>();

        /// <summary>
        /// A dictionary of sync states indexed by the values of distinct fields.
        /// </summary>
        private ConcurrentDictionary<string, SyncState> byDistinctFields = new ConcurrentDictionary<string, SyncState>();

        private SyncStateManager() { }


        /// <summary>
        /// This is part of an attempt to stop the 'do you want to save' popups; save
        /// everything we've touched, whether or not we've set anything on it.
        /// </summary>
        public void BruteForceSaveAll()
        {
            foreach (SyncState state in GetSynchronisedItems())
            {
                try
                {
                    string typeName = state.GetType().Name;

                    switch (typeName)
                    {
                        // TODO: there's almost certainly a cleaner and safer way of despatching this.
                        case "CallSyncState":
                        case "MeetingSyncState":
                            var outlookItem = (state as AppointmentSyncState).OutlookItem;
                            if (outlookItem.IsValid())
                            {
                                outlookItem.Save();
                            }
                            break;
                        case "ContactSyncState":
                            (state as ContactSyncState).OutlookItem.Save();
                            break;
                        case "TaskSyncState":
                            (state as TaskSyncState).OutlookItem.Save();
                            break;
                        default:
                            Globals.ThisAddIn.Log.AddEntry($"Unexpected type {typeName} in BruteForceSaveAll",
                                SuiteCRMClient.Logging.LogEntryType.Error);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Globals.ThisAddIn.Log.Warn("Exception during save all (probably not important)", ex);
                }
            }
        }

        /// <summary>
        /// Get all the syncstates I am holding.
        /// </summary>
        /// <returns>A collection of the items which I hold which are of the specified type.</returns>
        internal ICollection<SyncState> GetSynchronisedItems()
        {
            return this.byOutlookId.Values.ToList().AsReadOnly();
        }


        /// <summary>
        /// Get all the syncstates I am holding which are of this type.
        /// </summary>
        /// <typeparam name="SyncStateType">The type which is requested.</typeparam>
        /// <returns>A collection of the items which I hold which are of the specified type.</returns>
        internal ICollection<SyncStateType> GetSynchronisedItems<SyncStateType>() where SyncStateType : SyncState
        {
            return this.byOutlookId.Values.Select(x => x as SyncStateType).Where(x => x != null).ToList<SyncStateType>().AsReadOnly();
        }


        /// <summary>
        /// Count the number of items I monitor.
        /// </summary>
        /// <returns>A count of the number of items I monitor.</returns>
        public int CountItems()
        {
            return byOutlookId.Values.Count();
        }


        /// <summary>
        /// Get the existing sync state for this item, if it exists and is of the appropriate
        /// type, else null.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <returns>The appropriate sync state, or null if none.</returns>
        /// <exception cref="UnexpectedSyncStateClassException">if the sync state found is not of the expected class (shouldn't happen).</exception>
        public SyncState<ItemType> GetExistingSyncState<ItemType>(ItemType item)
            where ItemType : class
        {
            SyncState<ItemType> result;

            string typeName = Microsoft.VisualBasic.Information.TypeName(item);

            try
            {
                switch (typeName)
                {
                    // TODO: there's almost certainly a cleaner and safer way of despatching this.
                    case "AppointmentItem":
                        result = this.GetSyncState(item as Outlook.AppointmentItem) as SyncState<ItemType>;
                        break;
                    case "ContactItem":
                        result = this.GetSyncState(item as Outlook.ContactItem) as SyncState<ItemType>;
                        break;
                    case "TaskItem":
                        result = this.GetSyncState(item as Outlook.TaskItem) as SyncState<ItemType>;
                        break;
                    default:
                        Globals.ThisAddIn.Log.AddEntry($"Unexpected type {typeName} in GetExistingSyncState",
                            SuiteCRMClient.Logging.LogEntryType.Error);
                        result = null;
                        break;
                }
            }
            catch (TypeInitializationException tix)
            {
                log.Warn("Bad CRM id?", tix);
                result = null;
            }
            catch (KeyNotFoundException kex)
            {
                log.Warn("KeyNotFoundException in GetExistingSyncState", kex);
                result = null;
            }

            return result;
        }


        /// <summary>
        /// Get the existing sync state for this item, if it exists and is of the appropriate
        /// type, else null.
        /// </summary>
        /// <remarks>Outlook items are not true objects and don't have a common superclass, 
        /// so we have to use this rather clumsy overloading.</remarks>
        /// <param name="appointment">The item.</param>
        /// <returns>The appropriate sync state, or null if none.</returns>
        /// <exception cref="UnexpectedSyncStateClassException">if the sync state found is not of the expected class (shouldn't happen).</exception>
        public AppointmentSyncState GetSyncState(Outlook.AppointmentItem appointment)
        {
            SyncState result;

            try
            {
                result = (appointment.IsValid() && this.byOutlookId.ContainsKey(appointment.EntryID)) ? this.byOutlookId[appointment.EntryID] : null;
                CrmId crmId = result == null ? appointment.GetCrmId() : CheckForDuplicateSyncState(result, appointment.GetCrmId());

                if (CrmId.IsValid(crmId))
                {
                    if (result == null && this.byCrmId.ContainsKey(crmId))
                    {
                        result = this.byCrmId[crmId];
                    }
                    else if (result != null && this.byCrmId.ContainsKey(crmId) == false)
                    {
                        this.byCrmId[crmId] = result;
                        result.CrmEntryId = crmId;
                    }
                }

                if (result != null && !(result is AppointmentSyncState))
                {
                    throw new UnexpectedSyncStateClassException("AppointmentSyncState", result);
                }
            }
            catch (COMException)
            {
                // dead item passed.
                result = null;
            }

            return result as AppointmentSyncState;
        }


        /// <summary>
        /// Get the existing sync state for this item, if it exists and is of the appropriate
        /// type, else null.
        /// </summary>
        /// <remarks>Outlook items are not true objects and don't have a common superclass, 
        /// so we have to use this rather clumsy overloading.</remarks>
        /// <param name="contact">The item.</param>
        /// <returns>The appropriate sync state, or null if none.</returns>
        /// <exception cref="UnexpectedSyncStateClassException">if the sync state found is not of the expected class (shouldn't happen).</exception>
        public ContactSyncState GetSyncState(Outlook.ContactItem contact)
        {
            SyncState result;

            try
            {
                result = this.byOutlookId.ContainsKey(contact.EntryID) ? this.byOutlookId[contact.EntryID] : null;
                CrmId crmId = CheckForDuplicateSyncState(result, contact.GetCrmId());

                if (CrmId.IsValid(crmId))
                {
                    if (result == null && this.byCrmId.ContainsKey(crmId))
                    {
                        result = this.byCrmId[crmId];
                    }
                    else if (result != null && this.byCrmId.ContainsKey(crmId) == false)
                    {
                        this.byCrmId[crmId] = result;
                        result.CrmEntryId = crmId;
                    }
                }

                if (result != null && result as ContactSyncState == null)
                {
                    throw new UnexpectedSyncStateClassException("ContactSyncState", result);
                }
            }
            catch (COMException)
            {
                // dead item passed.
                result = null;
            }

            return result as ContactSyncState;
        }


        /// <summary>
        /// Get the existing sync state for this item, if it exists and is of the appropriate
        /// type, else null.
        /// </summary>
        /// <remarks>Outlook items are not true objects and don't have a common superclass, 
        /// so we have to use this rather clumsy overloading.</remarks>
        /// <param name="task">The item.</param>
        /// <returns>The appropriate sync state, or null if none.</returns>
        /// <exception cref="UnexpectedSyncStateClassException">if the sync state found is not of the expected class (shouldn't happen).</exception>
        public TaskSyncState GetSyncState(Outlook.TaskItem task)
        {
            SyncState result;

            try
            {
                result = this.byOutlookId.ContainsKey(task.EntryID) ? this.byOutlookId[task.EntryID] : null;
                CrmId crmId = result == null ? task.GetCrmId() : CheckForDuplicateSyncState(result, task.GetCrmId());

                if (CrmId.IsValid(crmId))
                {
                    if (result == null && this.byCrmId.ContainsKey(crmId))
                    {
                        result = this.byCrmId[crmId];
                    }
                    else if (result != null && crmId != null && this.byCrmId.ContainsKey(crmId) == false)
                    {
                        this.byCrmId[crmId] = result;
                        result.CrmEntryId = crmId;
                    }
                }

                if (result != null && !(result is TaskSyncState))
                {
                    throw new UnexpectedSyncStateClassException("TaskSyncState", result);
                }
            }
            catch (COMException)
            {
                // dead item passed.
                result = null;
            }

            return result as TaskSyncState;
        }


        /// <summary>
        /// Check whether there exists a sync state other than this state whose CRM id is 
        /// this CRM id or the CRM id of this state.
        /// </summary>
        /// <param name="state">The sync state to be checked.</param>
        /// <param name="crmId">A candidate CRM id.</param>
        /// <returns>A better guess at the CRM id.</returns>
        /// <exception cref="DuplicateSyncStateException">If a duplicate is detected.</exception>
        private CrmId CheckForDuplicateSyncState(SyncState state, CrmId crmId)
        {
            CrmId result = CrmId.IsInvalid(crmId) && state != null ? state.CrmEntryId : crmId;

            if (result != null)
            {
                SyncState byCrmState = this.byCrmId.ContainsKey(crmId) ? this.byCrmId[crmId] : null;

                if (state != null && byCrmState != null && state != byCrmState)
                {
                    throw new DuplicateSyncStateException(state);
                }
            }

            return result;
        }


        /// <summary>
        /// Get the existing sync state for this CRM item, if it exists, else null.
        /// </summary>
        /// <param name="crmItem">The item.</param>
        /// <returns>The appropriate sync state, or null if none.</returns>
        public SyncState GetExistingSyncState(EntryValue crmItem)
        {
            SyncState result;
            string outlookId = crmItem.GetValueAsString("outlook_id");
            CrmId crmId = CrmId.Get(crmItem.id);

            if (this.byCrmId.ContainsKey(crmId))
            {
                result = this.byCrmId[crmId];
            }
            else if (this.byOutlookId.ContainsKey(outlookId))
            {
                result = this.byOutlookId[outlookId];
            }
            else if (this.byGlobalId.ContainsKey(outlookId))
            {
                result = this.byGlobalId[outlookId];
            }
            else
            {
                string simulatedGlobalId = SyncStateManager.SimulateGlobalId(crmId);

                if (this.byGlobalId.ContainsKey(simulatedGlobalId))
                {
                    result = this.byGlobalId[simulatedGlobalId];
                }
                else
                {
                    string distinctFields = GetDistinctFields(crmItem);

                    if (string.IsNullOrEmpty(distinctFields))
                    {
                        result = null;
                    }
                    else if (this.byDistinctFields.ContainsKey(distinctFields))
                    {
                        result = this.byDistinctFields[distinctFields];
                    }
                    else
                    {
                        result = null;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Get the existing sync state for this CRM item, if it exists, else null.
        /// </summary>
        /// <param name="outlookId">The Outlook id of the syncstate to seek.</param>
        /// <param name="crmId">The CRM id of the syncstate to seek.</param>
         /// <returns>The appropriate sync state, or null if none.</returns>
        public SyncState GetExistingSyncState(string outlookId, CrmId crmId)
        {
            SyncState result;

            if (this.byCrmId.ContainsKey(crmId))
            {
                result = this.byCrmId[crmId];
            }
            else if (!string.IsNullOrEmpty(outlookId))
            {
                if (this.byOutlookId.ContainsKey(outlookId))
                {
                    result = this.byOutlookId[outlookId];
                }
                else
                {
                    result = this.byGlobalId.ContainsKey(outlookId) ?
                        this.byGlobalId[outlookId] :
                        null;
                }
            }
            else
            {
                string simulatedGlobalId = SyncStateManager.SimulateGlobalId(crmId);

                result = this.byGlobalId.ContainsKey(simulatedGlobalId) ? this.byGlobalId[simulatedGlobalId] : null;
            }

            return result;
        }


        /// <summary>
        /// Get a string representing the values of the distinct fields of this crmItem, 
        /// as a final fallback for identifying an otherwise unidentifiable object.
        /// </summary>
        /// <param name="crmItem">An item received from CRM.</param>
        /// <returns>An identifying string.</returns>
        /// <see cref="SyncState{ItemType}.IdentifyingFields"/> 
        private string GetDistinctFields(EntryValue crmItem)
        {
            string result;

            switch(crmItem.module_name)
            {
                case CallsSynchroniser.CrmModule:
                    result = CallSyncState.GetDistinctFields(crmItem);
                    break;
                case ContactSynchroniser.CrmModule:
                    result = ContactSyncState.GetDistinctFields(crmItem);
                    break;
                case MeetingsSynchroniser.CrmModule:
                    result = MeetingSyncState.GetDistinctFields(crmItem);
                    break;
                case TaskSynchroniser.CrmModule:
                    result = TaskSyncState.GetDistinctFields(crmItem);
                    break;
                default:
                    this.log.Warn($"Unexpected CRM module name '{crmItem.module_name}'");
                    result = string.Empty;
                    break;
            }

            return result;
        }


        /// <summary>
        /// Get a sync state for this item, creating it if necessary.
        /// </summary>
        /// <remarks>Outlook items are not true objects and don't have a common superclass, 
        /// so we have to use this rather clumsy despatch.</remarks>
        /// <param name="item">the item.</param>
        /// <returns>an appropriate sync state.</returns>
        /// <exception cref="UnexpectedSyncStateClassException">if the sync state found is not of the expected class (shouldn't happen).</exception>
        public SyncState<ItemType> GetOrCreateSyncState<ItemType>(ItemType item)
            where ItemType : class
        {
            lock (this.creationLock)
            {
                SyncState<ItemType> result = this.GetExistingSyncState(item);

                if (result == null)
                result = CreateSyncStateForItem(item);

                return result;
            }
        }

        private SyncState<ItemType> CreateSyncStateForItem<ItemType>(ItemType item)
            where ItemType : class
        {
            SyncState<ItemType> result;
            string outlookId;
            var typeName = Microsoft.VisualBasic.Information.TypeName(item);

            try
            {
                switch (typeName)
                {
                    // TODO: there's almost certainly a cleaner and safer way of despatching this.
                    case "AppointmentItem":
                        outlookId = ((Outlook.AppointmentItem)item).EntryID;
                        result = this.CreateSyncState(item as Outlook.AppointmentItem) as SyncState<ItemType>;
                        break;
                    case "ContactItem":
                        outlookId = ((Outlook.ContactItem)item).EntryID;
                        result = this.CreateSyncState(item as Outlook.ContactItem) as SyncState<ItemType>;
                        break;
                    case "TaskItem":
                        outlookId = ((Outlook.TaskItem)item).EntryID;
                        result = this.CreateSyncState(item as Outlook.TaskItem) as SyncState<ItemType>;
                        break;
                    default:
                        Globals.ThisAddIn.Log.AddEntry($"Unexpected type {typeName} in CreateSyncStateForItem",
                            SuiteCRMClient.Logging.LogEntryType.Error);
                        result = null;
                        break;
                }

                if (result != null)
                {
                    this.byDistinctFields[result.IdentifyingFields] = result;
                }
            }
            catch (COMException)
            {
                Globals.ThisAddIn.Log.Error("Invalid or detached COM object passed to CreateSyncStateForItem");
                result = null;
            }

            return result;
        }


        /// <summary>
        /// Create an appropriate sync state for an appointment item.
        /// </summary>
        /// <remarks>Outlook items are not true objects and don't have a common superclass, 
        /// so we have to use this rather clumsy overloading.</remarks>
        /// <param name="appointment">The item.</param>
        /// <returns>An appropriate sync state, or null if the appointment was invalid.</returns>
        private AppointmentSyncState CreateSyncState(Outlook.AppointmentItem appointment)
        {
            AppointmentSyncState result;

            CrmId crmId = appointment.GetCrmId();

            if (CrmId.IsValid(crmId) && this.byCrmId.ContainsKey(crmId) && this.byCrmId[crmId] != null)
            {
                result = CheckUnexpectedFoundState<Outlook.AppointmentItem, AppointmentSyncState>(appointment, crmId);
            }
            else
            {
                var modifiedDate = ParseDateTimeFromUserProperty(appointment.UserProperties[ModifiedDatePropertyName]);
                if (appointment.IsCall())
                {
                    result = this.SetByOutlookId<AppointmentSyncState>(appointment.EntryID,
                        new CallSyncState(appointment, crmId, modifiedDate));
                }
                else
                {
                    result = this.SetByOutlookId<AppointmentSyncState>(appointment.EntryID,
                        new MeetingSyncState(appointment, crmId, modifiedDate));
                }
                this.byGlobalId[appointment.GlobalAppointmentID] = result;
            }

            if (result != null && CrmId.IsValid(crmId))
            {
                this.byCrmId[crmId] = result;
            }

            return result;
        }

        /// <summary>
        /// Return the global appointment id corresponding to this crm id.
        /// </summary>
        /// <remarks>
        /// Outlook items arriving from CRM have a global appointment id based on their CRM id.
        /// Arguably this should not go here.
        /// </remarks>
        /// <see cref="Outlook.AppointmentItem.GlobalAppointmentId"/> 
        /// <see cref="AppointmentItemExtension.GetVCalId(Outlook.AppointmentItem)"/> 
        /// <param name="crmId">The CRM id from which the global id should be reverse engineered.</param>
        /// <returns>The reverse-engineered global id.</returns>
        private static string SimulateGlobalId(CrmId crmId)
        {
            byte[] globalId = new byte[89];
            byte[] header = new byte[40]
            {
                0x04, // 1: EOT
                0x00,
                0x00,
                0x00,
                0x82, // 5: left parenthesis
                0x00,
                0xE0, // 7: shift out
                0x00,
                0x74, // 9: G
                0xC5, // 10: right square bracket
                0xB7, // 11: left brace
                0x10, // 12: start of heading
                0x1A, // 13: ?
                0x82, // 14: left parenthesis
                0xE0, // 15: shift out
                0x08, // 16: ?
                0x00,
                0x00,
                0x00,
                0x00, // 20
                0x00,
                0x00,
                0x00,
                0x00,
                0x00, // 25
                0x00,
                0x00,
                0x00,
                0x00,
                0x00, // 30
                0x00,
                0x00,
                0x00,
                0x00,
                0x00, // 35
                0x00,
                0x31, // 37: s
                0x00,
                0x00,
                0x00  // 40
            };

            byte[] signature = Encoding.UTF8.GetBytes("vCal-Uid");
            byte[] crmIdBytes = Encoding.UTF8.GetBytes(crmId.ToString());

            Buffer.BlockCopy(header, 0, globalId, 0, 40);
            Buffer.BlockCopy(signature, 0, globalId, 40, signature.Length);
            int cursor = 40 + signature.Length;
            globalId[cursor++] = 0x01;
            globalId[cursor++] = 0;
            globalId[cursor++] = 0;
            Buffer.BlockCopy(crmIdBytes, 0, globalId, cursor, crmIdBytes.Length);

            string result = Encoding.UTF8.GetString(globalId, 0, globalId.Length);

            return result;
        }


        /// <summary>
        /// Create an appropriate sync state for an contact item.
        /// </summary>
        /// <remarks>Outlook items are not true objects and don't have a common superclass, 
        /// so we have to use this rather clumsy overloading.</remarks>
        /// <param name="contact">The item.</param>
        /// <returns>An appropriate sync state, or null if the contact was invalid.</returns>
        private ContactSyncState CreateSyncState(Outlook.ContactItem contact)
        {
            ContactSyncState result;

            CrmId crmId = contact.GetCrmId();
            if (CrmId.IsValid(crmId) && this.byCrmId.ContainsKey(crmId) && this.byCrmId[crmId] != null)
            {
                result = CheckUnexpectedFoundState<Outlook.ContactItem, ContactSyncState>(contact, crmId);
            }
            else
            {
                result = this.SetByOutlookId<ContactSyncState>(contact.EntryID,
                    new ContactSyncState(contact, crmId,
                        ParseDateTimeFromUserProperty(contact.UserProperties[ModifiedDatePropertyName])));
            }

            if (result != null && CrmId.IsValid(crmId))
            {
                this.byCrmId[crmId] = result;
            }

            return result;
        }


        /// <summary>
        /// Create an appropriate sync state for an task item.
        /// </summary>
        /// <remarks>Outlook items are not true objects and don't have a common superclass, 
        /// so we have to use this rather clumsy overloading.</remarks>
        /// <param name="task">The item.</param>
        /// <returns>An appropriate sync state, or null if the task was invalid.</returns>
        private TaskSyncState CreateSyncState(Outlook.TaskItem task)
        {
            TaskSyncState result;

            CrmId crmId = task.GetCrmId();

            if (CrmId.IsValid(crmId) && this.byCrmId.ContainsKey(crmId) && this.byCrmId[crmId] != null)
            {
                result = CheckUnexpectedFoundState<Outlook.TaskItem, TaskSyncState>(task, crmId);
            }
            else
            {
                result = this.SetByOutlookId<TaskSyncState>(task.EntryID,
                    new TaskSyncState(task, crmId,
                        ParseDateTimeFromUserProperty(task.UserProperties[ModifiedDatePropertyName])));
            }

            if (result != null && CrmId.IsValid(crmId))
            {
                this.byCrmId[crmId] = result;
            }

            return result;
        }


        /// <summary>
        /// Perform sanity checks on an unexpected sync state has been found where we expected 
        /// to find none. 
        /// </summary>
        /// <remarks>This probably shouldn't happen and perhaps ought to be flagged as a hard error
        /// anyway.</remarks>
        /// <typeparam name="ItemType">The type of Outlook item being considered.</typeparam>
        /// <typeparam name="StateType">The appropriate sync state class for that item.</typeparam>
        /// <param name="olItem">The Outlook item being considered.</param>
        /// <param name="crmId">The CRM id associated with that item.</param>
        /// <returns>An appropriate sync state</returns>
        /// <exception cref="Exception">If the sync state doesn't exactly match what we would expect.</exception>
        private StateType CheckUnexpectedFoundState<ItemType, StateType>(ItemType olItem, CrmId crmId)
            where ItemType : class
            where StateType : SyncState<ItemType>
        {
            StateType result;
            var state = this.byCrmId.ContainsKey(crmId) ? this.byCrmId[crmId] : null;

            if (state != null)
            {
                result = state as StateType;

                if (result == null)
                {
                    throw new Exception($"Unexpected state type found: {state.GetType().Name}.");
                }
                else if (!result.OutlookItem.Equals(olItem))
                {
                    throw new ProbableDuplicateItemException<ItemType>(olItem, $"Probable duplicate Outlook item; crmId is {crmId}; identifying fields are {result.IdentifyingFields}");
                }
            }
            else
            {
                result = this.CreateSyncStateForItem(olItem) as StateType;
            }

            return result;
        }


        /// <summary>
        /// Set the value of my <see cref="byOutlookId"/> dictionary for this key to this value, 
        /// provided it is not already set.
        /// </summary>
        /// <remarks>
        /// This is part of defence against ending up with duplicate SyncStates, which is an 
        /// exceedingly bad thing. Values should not be put into the <see cref="byOutlookId"/> 
        /// dictionary except through this method. 
        /// </remarks>
        /// <typeparam name="StateType">The type of <see cref="SyncState"/> I am passing.</typeparam>
        /// <param name="key">The key I am setting.</param>
        /// <param name="value">The value I am seeking to set it to.</param>
        /// <returns>The value that is set.</returns>
        private StateType SetByOutlookId<StateType>(string key, StateType value)
            where StateType : SyncState
        {
            StateType result;

            if (!string.IsNullOrEmpty(key))
            {
                try
                {
                    var current = this.byOutlookId[key];
                    result = current as StateType;
                }
                catch (KeyNotFoundException)
                {
                    this.byOutlookId[key] = value;
                    result = value;
                }
            }
            else
            {
                result = null;
            }

            return result;
        }


        /// <summary>
        /// Parse a date/time from the supplied user property, if any.
        /// </summary>
        /// <param name="property">The property (may be null)</param>
        /// <returns>A date/time.</returns>
        private DateTime ParseDateTimeFromUserProperty(Outlook.UserProperty property)
        {
            DateTime result;
            if (property == null || property.Value == null)
            {
                result = default(DateTime);
            }
            else
            {
                result = DateTime.UtcNow;
                if (!DateTime.TryParseExact(property.Value, "yyyy-MM-dd HH:mm:ss", null, DateTimeStyles.None, out result))
                {
                    DateTime.TryParse(property.Value, out result);
                }
            }

            return result;
        }


        internal void RemoveOutlookId(string outlookItemEntryId)
        {
            this.byOutlookId[outlookItemEntryId] = null;
        }


        /// <summary>
        /// Remove all references to this sync state, if I hold any.
        /// </summary>
        /// <param name="state">The state to remove.</param>
        internal void RemoveSyncState(SyncState state)
        {
            lock (this.creationLock)
            {
                SyncState ignore;
                try
                {
                    if (this.byOutlookId[state.OutlookItemEntryId] == state)
                    {
                        this.byOutlookId.TryRemove(state.OutlookItemEntryId, out ignore);
                    }
                }
                catch (KeyNotFoundException) { }
                catch (COMException) { }
                try
                {
                    if (CrmId.IsValid(state.CrmEntryId) && this.byCrmId[state.CrmEntryId] == state)
                    {
                        this.byCrmId.TryRemove(state.CrmEntryId, out ignore);
                    }
                }
                catch (KeyNotFoundException) { }
            }
        }


        /// <summary>
        /// Called after a new item has been created from a CRM item, to fix up the index and prevent duplication.
        /// </summary>
        /// <typeparam name="SyncStateType">The type of <see cref="SyncState"/> passed</typeparam>
        /// <param name="crmId">The crmId to index this sync state to.</param>
        /// <param name="syncState">the sync state to index.</param>
        internal void SetByCrmId<SyncStateType>(CrmId crmId, SyncStateType syncState) where SyncStateType : SyncState
        {
            if (this.byCrmId.ContainsKey(crmId) && this.byCrmId[crmId] != null && this.byCrmId[crmId] != syncState)
            {
                throw new DuplicateCrmIdException(syncState, crmId);
            }

            this.byCrmId[crmId] = syncState;
            string outlookId = syncState?.OutlookItemEntryId;

            if (!string.IsNullOrEmpty(outlookId))
            {
                this.byOutlookId[outlookId] = syncState;
            }
        }
    }
}
