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
    using Extensions;
    using ProtoItems;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using System.Text;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// An agent which synchronises Outlook Task items with CRM.
    /// </summary>
    public class TaskSynchroniser: Synchroniser<Outlook.TaskItem, TaskSyncState>
    {
        /// <summary>
        /// The module I synchronise with.
        /// </summary>
        public const string CrmModule = "Tasks";

        public TaskSynchroniser(string name, SyncContext context)
            : base(name, context)
        {
            this.fetchQueryPrefix = string.Empty;
        }


        public override string DefaultCrmModule
        {
            get
            {
                return TaskSynchroniser.CrmModule;
            }
        }

        public override SyncDirection.Direction Direction => Properties.Settings.Default.SyncTasks;

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
            try
            {
                olItem?.Save();
                LogItemAction(olItem, "TaskSyncing.SaveItem, saved item");
            }
            catch (System.Exception any)
            {
                ErrorHandler.Handle($"Failure while saving task {olItem?.Subject}", any);
            }
        }

        /// <summary>
        /// Synchronise items in the specified folder with the specified SuiteCRM module.
        /// </summary>
        /// <param name="folder">The folder.</param>
        /// <param name="crmModule">The module to snychronise with.</param>
        protected override void SyncFolder(Outlook.MAPIFolder folder, string crmModule)
        {
            Log.Info($"TaskSyncing.SyncFolder: '{folder.FolderPath}'");
            try
            {
                var untouched =  new HashSet<SyncState<Outlook.TaskItem>>(SyncStateManager.Instance.GetSynchronisedItems<SyncState<Outlook.TaskItem>>());

                IList<EntryValue> records = MergeRecordsFromCrm(folder, crmModule, untouched);
                this.AddOrUpdateItemsFromCrmToOutlook(records, folder, untouched, crmModule);

                try
                {
                    ResolveUnmatchedItems(untouched);
                }
                catch (Exception ex)
                {
                    ErrorHandler.Handle("Failure while synchronising Tasks", ex);
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle("Failure while synchronising Tasks", ex);
            }
        }


        /// <summary>
        /// Log a message regarding this Outlook appointment.
        /// </summary>
        /// <param name="olItem">The outlook item.</param>
        /// <param name="message">The message to be logged.</param>
        internal override void LogItemAction(Outlook.TaskItem olItem, string message)
        {
            if (olItem != null && olItem.IsValid())
            {
                try
                {
                    CrmId crmId = this.IsEnabled() ? olItem.GetCrmId() : CrmId.Empty;
                    if (CrmId.IsInvalid(crmId)) { crmId = CrmId.Empty; }

                    StringBuilder bob = new StringBuilder();
                    bob.Append($"{message}:\n\tOutlook Id  : {olItem.EntryID}")
                        .Append(this.IsEnabled() ? $"\n\tCRM Id      : {crmId}" : string.Empty)
                        .Append($"\n\tSubject     : '{olItem.Subject}'")
                        .Append($"\n\tStatus      : {olItem.Status}")
                        .Append($"\n\tSensitivity : {olItem.Sensitivity}")
                        .Append($"\n\tTxState     : {SyncStateManager.Instance.GetExistingSyncState(olItem)?.TxState}");
                    Log.Info(bob.ToString());
                }
                catch (COMException)
                {
                    // Ignore: happens if the outlook item is already deleted.
                }
            }
        }


        protected override bool ShouldAddOrUpdateItemFromCrmToOutlook(Outlook.MAPIFolder folder, string crmType, EntryValue crmItem)
        {
            try
            {
                return RestAPIWrapper.GetUserId().Equals(crmItem.GetValueAsString("assigned_user_id"));
            }
            catch (TypeInitializationException tix)
            {
                Log.Warn("Bad CRM id?", tix);
                return false;
            }
        }


        protected override SyncState<Outlook.TaskItem> AddOrUpdateItemFromCrmToOutlook(Outlook.MAPIFolder tasksFolder,
            string crmType, EntryValue crmItem)
        {
            SyncState<Outlook.TaskItem> result = null;

            Log.Debug(
                $"TaskSyncing.AddOrUpdateItemFromCrmToOutlook\n\tSubject: {crmItem.GetValueAsString("name")}\n\tCurrent user id {RestAPIWrapper.GetUserId()}\n\tAssigned user id: {crmItem.GetValueAsString("assigned_user_id")}");

            var syncState = SyncStateManager.Instance.GetExistingSyncState(crmItem) as SyncState<Outlook.TaskItem>;

            result = syncState == null ? MaybeAddNewItemFromCrmToOutlook(tasksFolder, crmItem) : UpdateExistingOutlookItemFromCrm(crmItem, syncState);

            return result;
        }

        /// <summary>
        /// Item creation really ought to happen within the context of a lock, in order to prevent duplicate creation.
        /// </summary>
        /// <param name="tasksFolder">The folder in which the item should be created.</param>
        /// <param name="crmItem">The CRM item it will represent.</param>
        /// <returns>A syncstate whose Outlook item is the Outlook item representing this crmItem.</returns>
        private SyncState<Outlook.TaskItem> MaybeAddNewItemFromCrmToOutlook(Outlook.MAPIFolder tasksFolder, EntryValue crmItem)
        {
            SyncState<Outlook.TaskItem> result;

            lock (creationLock)
            {
                /* check for howlaround */
                var matches = this.FindMatches(crmItem);

                if (matches.Count == 0)
                {
                    /* didn't find it, so add it to Outlook */
                    result = AddNewItemFromCrmToOutlook(tasksFolder, crmItem);
                }
                else
                {
                    this.Log.Warn($"Howlaround detected? Task '{crmItem.GetValueAsString("name")}' offered with id {crmItem.GetValueAsString("id")}, expected {matches[0].CrmEntryId}, {matches.Count} duplicates");
                    result = matches[0];
                }
            }

            return result;
        }

        private static string ExtractTime(DateTime dateStart)
        {
            return
                                TimeSpan.FromHours(dateStart.Hour)
                                    .Add(TimeSpan.FromMinutes(dateStart.Minute))
                                    .ToString(@"hh\:mm");
        }

        private SyncState<Outlook.TaskItem> UpdateExistingOutlookItemFromCrm(EntryValue crmItem, SyncState<Outlook.TaskItem> syncState)
        {
            if (!syncState.IsDeletedInOutlook)
            {
                Outlook.TaskItem olItem = syncState.OutlookItem;

                if (olItem.IsValid())
                {
                    Outlook.UserProperty oProp = olItem.UserProperties[SyncStateManager.ModifiedDatePropertyName];

                    if (oProp.Value != crmItem.GetValueAsString("date_modified"))
                    {
                        SetOutlookItemPropertiesFromCrmItem(crmItem, olItem);
                    }
                    syncState.OModifiedDate = DateTime.ParseExact(crmItem.GetValueAsString("date_modified"), "yyyy-MM-dd HH:mm:ss", null);
                }
                else
                {
                    Log.Error($"Attempting to update invalid Outlook item '{crmItem.GetValueAsString("name")}'");
                }
            }
            return syncState;
        }

        private void SetOutlookItemPropertiesFromCrmItem(EntryValue crmItem, Outlook.TaskItem olItem)
        {
            try
            {
                DateTime dateStart = crmItem.GetValueAsDateTime("date_start");
                DateTime dateDue = crmItem.GetValueAsDateTime("date_due");
                string timeStart = ExtractTime(dateStart);
                string timeDue = ExtractTime(dateDue);

                olItem.Subject = crmItem.GetValueAsString("name");

                try
                {
                    olItem.StartDate = MaybeChangeDate(dateStart, olItem.StartDate, "syncState.StartDate");
                }
                catch (COMException comx)
                {
#if DEBUG
                    Log.Debug($"COM Exception while trying to set start date of task: '{comx.Message}'. Some otherwise-valid tasks don't support this");
#endif
                }

                try { 
                olItem.DueDate = MaybeChangeDate(dateDue, olItem.DueDate, "syncState.DueDate");
                }
                catch (COMException comx)
                {
#if DEBUG
                    Log.Debug($"COM Exception while trying to set start date of task: '{comx.Message}'. Do some otherwise-valid tasks not support this?");
#endif
                }

                string body = crmItem.GetValueAsString("description");
                olItem.Body = string.Concat(body, "#<", timeStart, "#", timeDue);
                olItem.Status = GetStatus(crmItem.GetValueAsString("status"));
                olItem.Importance = GetImportance(crmItem.GetValueAsString("priority"));
                EnsureSynchronisationPropertiesForOutlookItem(olItem, crmItem.GetValueAsString("date_modified"), DefaultCrmModule, CrmId.Get(crmItem.id));
            }
            finally
            {
                this.SaveItem(olItem);
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
                if (newValue != null && newValue.HasValue && newValue.Value > DateTime.MinValue)
                {
                    Log.Info($"\tt{nameOfValue}= {oldValue}, newValue= {newValue.Value}");
                    result = newValue.Value;
                }
            }
            catch (Exception fail)
            {
                /* you (sometimes? always?) can't set the start or due dates of tasks. Investigate. */
                ErrorHandler.Handle(
                    $"Failed to set {nameOfValue} on task",
                    fail);
            }

            return result;
        }

        private SyncState<Outlook.TaskItem> AddNewItemFromCrmToOutlook(Outlook.MAPIFolder tasksFolder, EntryValue crmItem)
        {
            TaskSyncState result = null;
            Log.Debug(
                (string)string.Format(
                    $"{this.GetType().Name}.AddNewItemFromCrmToOutlook, entry id is '{crmItem.GetValueAsString("id")}', creating in Outlook."));

            /*
             * There's a nasty little bug (#223) where Outlook offers us back in a different thread
             * the item we're creating, before we're able to set up the sync state which marks it
             * as already known. By locking on the enqueueing lock here, we should prevent that.
             */
            lock (enqueueingLock)
            {
                Outlook.TaskItem olItem = tasksFolder.Items.Add(Outlook.OlItemType.olTaskItem);

                if (olItem != null)
                {
                    try
                    {
                        this.SetOutlookItemPropertiesFromCrmItem(crmItem, olItem);
                    }
                    finally
                    {
                        result = SyncStateManager.Instance.GetOrCreateSyncState(olItem) as TaskSyncState;
                        result.SetNewFromCRM();

                        this.SaveItem(olItem);
                    }
                }
            }

            return result;
        }


        /// <summary>
        /// Construct a JSON packet representing the Outlook item of this sync state, and despatch 
        /// it to CRM.
        /// </summary>
        /// <param name="syncState">The Outlook item.</param>
        /// <returns>The CRM id of the object created or modified.</returns>
        protected override CrmId ConstructAndDespatchCrmItem(SyncState<Outlook.TaskItem> syncState)
        {
            return CrmId.Get(RestAPIWrapper.SetEntry(new ProtoTask(syncState.OutlookItem).AsNameValues(), this.DefaultCrmModule));
        }


        protected override void LinkOutlookItems(Outlook.MAPIFolder taskFolder)
        {
            try
            {
                Outlook.Items items = taskFolder.Items; //.Restrict("[MessageClass] = 'IPM.Task'" + GetStartDateString());
                foreach (Outlook.TaskItem olItem in items)
                {
                    if (olItem.DueDate >= DateTime.Now.AddDays(-5))
                    {
                        SyncStateManager.Instance.GetOrCreateSyncState(olItem).SetPresentAtStartup();
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle("Failed while trying to index Tasks", ex);
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
                if (!olProperty.Value.Equals(value))
                {
                    try
                    {
                        olProperty.Value = value ?? string.Empty;
                        Log.Debug($"TaskSyncing.EnsureSynchronisationPropertyForOutlookItem: Set property {name} to value {value} on item {olItem.Subject}");
                    }
                    finally
                    {
                        this.SaveItem(olItem);
                    }
                }
            }
            catch (Exception any)
            {
                ErrorHandler.Handle($"Failed to set property {name} to value {value} on task {olItem.Subject}", any);
            }
        }


        public override Outlook.MAPIFolder GetDefaultFolder()
        {
            return Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks);
        }

        internal override string GetOutlookEntryId(Outlook.TaskItem olItem)
        {
            return olItem.EntryID;
        }

        protected override CrmId GetCrmEntryId(Outlook.TaskItem olItem)
        {
            return olItem.GetCrmId();
        }

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
