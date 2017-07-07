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
    using SuiteCRMAddIn.ProtoItems;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// An agent which synchronises Outlook Task items with CRM.
    /// </summary>
    public class TaskSyncing: Synchroniser<Outlook.TaskItem>
    {
        /// <summary>
        /// The module I synchronise with.
        /// </summary>
        public const string CrmModule = "Tasks";

        public TaskSyncing(string name, SyncContext context)
            : base(name, context)
        {
            this.fetchQueryPrefix = string.Empty;
        }

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
                return TaskSyncing.CrmModule;
            }
        }

        public override SyncDirection.Direction Direction => Properties.Settings.Default.SyncCalendar;

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

        protected override void SaveItem(Outlook.TaskItem olItem)
        {
            olItem.Save();
            LogItemAction(olItem, "TaskSyncing.SaveItem, saved item");
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
                var untouched = new HashSet<SyncState<Outlook.TaskItem>>(ItemsSyncState);

                MergeRecordsFromCrm(folder, crmModule, untouched);

                try
                {
                    ResolveUnmatchedItems(untouched);
                }
                catch (Exception ex)
                {
                    Log.Error("TaskSyncing.SyncFolder", ex);
                }
            }
            catch (Exception ex)
            {
                Log.Error("TaskSyncing.SyncFolder", ex);
            }
        }


        /// <summary>
        /// Log a message regarding this Outlook appointment.
        /// </summary>
        /// <param name="olItem">The outlook item.</param>
        /// <param name="message">The message to be logged.</param>
        protected override void LogItemAction(Outlook.TaskItem olItem, string message)
        {
            try
            {
                Outlook.UserProperty olPropertyEntryId = olItem.UserProperties["SEntryID"];
                string crmId = olPropertyEntryId == null ?
                    "[not present]" :
                    olPropertyEntryId.Value;
                Log.Info($"{0}:\n\tOutlook Id  : {olItem.EntryID}\n\tCRM Id      : {crmId}\n\tSubject    : '{olItem.Subject}'\n\tStatus      : {olItem.Status}");
            }
            catch (COMException)
            {
                // Ignore: happens if the outlook item is already deleted.
            }
        }

        // TODO: this is very horrible and should be reworked.
        protected override SyncState<Outlook.TaskItem> AddOrUpdateItemFromCrmToOutlook(Outlook.MAPIFolder tasksFolder, string crmType, EntryValue crmItem)
        {
            SyncState<Outlook.TaskItem> result = null;

            Log.Debug($"TaskSyncing.AddOrUpdateItemFromCrmToOutlook\n\tSubject: {crmItem.GetValueAsString("name")}\n\tCurrent user id {RestAPIWrapper.GetUserId()}\n\tAssigned user id: {crmItem.GetValueAsString("assigned_user_id")}");

            if (RestAPIWrapper.GetUserId() == crmItem.GetValueAsString("assigned_user_id"))
            {
                DateTime dateStart = crmItem.GetValueAsDateTime("newValue");
                DateTime dateDue = crmItem.GetValueAsDateTime("dateDue");
                string timeStart =
                        TimeSpan.FromHours(dateStart.Hour)
                            .Add(TimeSpan.FromMinutes(dateStart.Minute))
                            .ToString(@"hh\:mm");
                string timeDue = TimeSpan.FromHours(dateDue.Hour)
                                .Add(TimeSpan.FromMinutes(dateDue.Minute))
                                .ToString(@"hh\:mm");

                var syncState = this.GetExistingSyncState(crmItem);

                if (syncState == null)
                {
                    /* check for howlaround */
                    var matches = this.FindMatches(crmItem);

                    if (matches.Count == 0)
                    {
                        /* didn't find it, so add it to Outlook */
                        result = AddNewItemFromCrmToOutlook(tasksFolder, crmItem, dateStart, dateDue, timeStart, timeDue);
                    }
                    else
                    {
                        this.Log.Warn($"Howlaround detected? Task '{crmItem.GetValueAsString("name")}' offered with id {crmItem.GetValueAsString("id")}, expected {matches[0].CrmEntryId}, {matches.Count} duplicates");
                    }
                }
                else
                {
                    result = UpdateExistingOutlookItemFromCrm(crmItem, dateStart, dateDue, timeStart, timeDue, syncState);
                }
            }

            return result;
        }

        private SyncState<Outlook.TaskItem> UpdateExistingOutlookItemFromCrm(EntryValue crmItem, DateTime? date_start, DateTime? date_due, string time_start, string time_due, SyncState<Outlook.TaskItem> syncStateForItem)
        {
            if (!syncStateForItem.IsDeletedInOutlook)
            {
                Outlook.TaskItem olItem = syncStateForItem.OutlookItem;
                Outlook.UserProperty oProp = olItem.UserProperties["SOModifiedDate"];

                if (oProp.Value != crmItem.GetValueAsString("date_modified"))
                {
                    SetOutlookItemPropertiesFromCrmItem(crmItem, date_start, date_due, time_start, time_due, olItem);
                }
                syncStateForItem.OModifiedDate = DateTime.ParseExact(crmItem.GetValueAsString("date_modified"), "yyyy-MM-dd HH:mm:ss", null);
            }
            return syncStateForItem;
        }

        private void SetOutlookItemPropertiesFromCrmItem(EntryValue crmItem, DateTime? dateStart, DateTime? dateDue, string timeStart, string timeDue, Outlook.TaskItem olItem)
        {
            try
            {
                olItem.Subject = crmItem.GetValueAsString("name");

                olItem.StartDate = MaybeChangeDate(dateStart, olItem.StartDate, "olItem.StartDate");

                olItem.DueDate = MaybeChangeDate(dateDue, olItem.DueDate, "olItem.DueDate");

                string body = crmItem.GetValueAsString("description");
                olItem.Body = string.Concat(body, "#<", timeStart, "#", timeDue);
                olItem.Status = GetStatus(crmItem.GetValueAsString("status"));
                olItem.Importance = GetImportance(crmItem.GetValueAsString("priority"));
                EnsureSynchronisationPropertiesForOutlookItem(olItem, crmItem.GetValueAsString("date_modified"), DefaultCrmModule, crmItem.id);
            }
            finally
            {
                olItem.Save();
            }
        }

        /// <summary>
        /// Return a DateTime which is this new value if the new value is valid, else this old value.
        /// </summary>
        /// <param name="newValue">A new value, which may be invalid (null or equal to DateTime.MinValue).</param>
        /// <param name="oldValue">An old value.</param>
        /// <param name="nameOfValue">The name of the value to be changed, for logging purposes.</param>
        /// <returns>a DateTime which is this new value if the new value is valid, else this old value.</returns>
        private DateTime MaybeChangeDate(DateTime? newValue, DateTime oldValue, string nameOfValue)
        {
            DateTime result = oldValue;
            try
            {
                if (newValue.HasValue && newValue.Value > DateTime.MinValue)
                {
                    Log.Warn($"\tt{nameOfValue}= {oldValue}, newValue= {newValue.Value}");
                    result = newValue.Value;
                }
            }
            catch (Exception fail)
            {
                /* you (sometimes? always?) can't set the start or due dates of tasks. Investigate. */
                Log.Error(
                    $"TaskSyncing.SetOutlookItemPropertiesFromCrmItem: Failed to set {nameOfValue} on task because {fail.Message}",
                    fail);
            }

            return result;
        }

        private SyncState<Outlook.TaskItem> AddNewItemFromCrmToOutlook(Outlook.MAPIFolder tasksFolder, EntryValue crmItem, DateTime? date_start, DateTime? date_due, string time_start, string time_due)
        {
            Outlook.TaskItem olItem = tasksFolder.Items.Add(Outlook.OlItemType.olTaskItem);
            TaskSyncState newState = null;

            try
            {
                this.SetOutlookItemPropertiesFromCrmItem(crmItem, date_start, date_due, time_start, time_due, olItem);

                newState = new TaskSyncState
                {
                    OutlookItem = olItem,
                    OModifiedDate = DateTime.ParseExact(crmItem.GetValueAsString("date_modified"), "yyyy-MM-dd HH:mm:ss", null),
                    CrmEntryId = crmItem.GetValueAsString("id"),
                };
                ItemsSyncState.Add(newState);
            }
            finally
            {
                olItem.Save();
                LogItemAction(olItem, "TaskSyncing.AddNewItemFromCrmToOutlook, saved item");
            }

            return newState;
        }


        /// <summary>
        /// Construct a JSON packet representing this Outlook item, and despatch it to CRM.
        /// </summary>
        /// <param name="olItem">The Outlook item.</param>
        /// <param name="crmType">The type within CRM to which the item should be added.</param>
        /// <param name="entryId">The corresponding entry id in CRM, if known.</param>
        /// <returns>The CRM id of the object created or modified.</returns>
        protected override string ConstructAndDespatchCrmItem(Outlook.TaskItem olItem, string crmType, string entryId)
        {
            return RestAPIWrapper.SetEntryUnsafe(new ProtoTask(olItem).AsNameValues(entryId), crmType);
        }


        protected override void GetOutlookItems(Outlook.MAPIFolder taskFolder)
        {
            try
            {
                Outlook.Items items = taskFolder.Items; //.Restrict("[MessageClass] = 'IPM.Task'" + GetStartDateString());
                foreach (Outlook.TaskItem oItem in items)
                {
                    if (oItem.DueDate < DateTime.Now.AddDays(-5))
                        continue;
                    Outlook.UserProperty oProp = oItem.UserProperties["SOModifiedDate"];
                    if (oProp != null)
                    {
                        Outlook.UserProperty oProp2 = oItem.UserProperties["SEntryID"];
                        ItemsSyncState.Add(new TaskSyncState
                        {
                            OutlookItem = oItem,
                            //OModifiedDate = "Fresh",
                            OModifiedDate = DateTime.UtcNow,

                            CrmEntryId = oProp2.Value.ToString()
                        });
                    }
                    else
                    {
                        ItemsSyncState.Add(new TaskSyncState
                        {
                            OutlookItem = oItem
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.GetOutlookTItems", ex);
            }
        }


        protected override void EnsureSynchronisationPropertyForOutlookItem(Outlook.TaskItem olItem, string name, string value)
        {
            try
            {
                Outlook.UserProperty olProperty = olItem.UserProperties[name];
                if (olProperty == null)
                {
                    olProperty = olItem.UserProperties.Add(name, Outlook.OlUserPropertyType.olText);
                }
                olProperty.Value = value ?? string.Empty;
            }
            finally
            {
                this.SaveItem(olItem);
            }
        }


        public override Outlook.MAPIFolder GetDefaultFolder()
        {
            return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks);
        }

        protected override SyncState<Outlook.TaskItem> ConstructSyncState(Outlook.TaskItem olItem)
        {
            return new TaskSyncState
            {
                OutlookItem = olItem,
                CrmEntryId = olItem.UserProperties["SEntryID"]?.Value.ToString(),
                OModifiedDate = ParseDateTimeFromUserProperty(olItem.UserProperties["SOModifiedDate"]?.Value.ToString()),
            };
        }

        internal override string GetOutlookEntryId(Outlook.TaskItem olItem)
        {
            return olItem.EntryID;
        }

        protected override bool IsCurrentView => Context.CurrentFolderItemType == Outlook.OlItemType.olTaskItem;

        // Should presumably be removed at some point. Existing code was ignoring deletions for Contacts and Tasks
        // (but not for Appointments).
        protected override bool PropagatesLocalDeletions => true;

        /// <summary>
        /// Return the sensitivity of this outlook item.
        /// </summary>
        /// <remarks>
        /// Outlook item classes do not inherit from a common base class, so generic client code cannot refer to 'OutlookItem.Sensitivity'.
        /// </remarks>
        /// <param name="olItem">The outlook item whose sensitivity is required.</param>
        /// <returns>the sensitivity of the item.</returns>
        internal override Outlook.OlSensitivity GetSensitivity(Outlook.TaskItem olItem)
        {
            return olItem.Sensitivity;
        }

        protected override bool IsMatch(Outlook.TaskItem olItem, EntryValue crmItem)
        {
            return olItem.Subject == crmItem.GetValueAsString("name") &&
                olItem.StartDate.ToUniversalTime() == crmItem.GetValueAsDateTime("newValue").ToUniversalTime();
        }
    }
}
