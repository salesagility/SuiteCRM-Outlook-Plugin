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
using System;
using System.Collections.Generic;

namespace SuiteCRMClient
{
    using RESTObjects;
    using System.Collections;
    using System.Linq;
    using Logging;

    /// <summary>
    /// A class which comprises wrappers for calls in the REST API, which return objects
    /// generally from the RESTObjects package.
    /// </summary>
    public static class RestAPIWrapper
    {
        private static ILogger Log;

        /// <summary>
        /// The list of the modules and their permissions change extremely rarely; 
        /// they may safely be cached for a session.
        /// </summary>
        private static AvailableModules modulesCache = null;

        /// <summary>
        /// A map that maps module names to the list of fields in the named module.
        /// </summary>
        /// <remarks>
        /// Module fields change equally rarely. Cache em, too!
        /// </remarks>
        private static Dictionary<string, ModuleFields> moduleFieldsCache = new Dictionary<string, ModuleFields>();

        public static UserSession SuiteCRMUserSession;

        /// <summary>
        /// We don't always want to call out to the server to get the user id, and sometimes 
        /// we really really don't. It's unlikely to change often in a session - in fact it 
        /// can currently change only when the user changes settings.
        /// </summary>
        public static string CachedUserId { get; private set; } = string.Empty;

        public static void SetLog(ILogger log)
        {
            Log = log;
        }

        /// <summary>
        /// Get the list of modules installed in the connected CRM instance, with their
        /// associated access control lists.
        /// </summary>
        /// <remarks>This data changes only rarely, and is consequently cached for the session.
        /// </remarks>
        /// <returns>the list of modules installed in the connected CRM instance.</returns>
        public static AvailableModules GetModules()
        {
            if (modulesCache == null)
            {
                EnsureLoggedIn();
                object data = new
                {
                    @session = SuiteCRMUserSession.id
                };
                try
                {
                    Log.Debug("Calling get_available_modules...");
                    modulesCache = SuiteCRMUserSession.RestServer.GetCrmResponse<AvailableModules>("get_available_modules", data);
                    Log.Debug("Successfully called get_available_modules.");
                }
                catch (Exception any)
                {
                    Log.Error($"Call to get_available_modules failed: {any.Message}");
                    throw;
                }
            }
            return modulesCache;             
        }

        /// <summary>
        /// Return only those modules which have relationships to the email module.
        /// </summary>
        /// <returns>A list of only those modules which have relationships to the email module.</returns>
        public static List<AvailableModule> GetModulesHavingEmailRelationships()
        {
            List<AvailableModule> modules = new List<AvailableModule>();
            foreach(AvailableModule module in GetModules().items)
            {
                try
                {
                    foreach (string field in GetFields(module.module_key))
                    {
                        if (field.StartsWith("email_") || field.EndsWith("_email"))
                        {
                            modules.Add(module);
                            break;
                        }
                    }
                }
                catch (Exception)
                {
                    Log.Debug($"SuiteCMHelper.GetModulesHavingEmailRelationships: failed to fetch fields list for {module.module_key}");
                }
            }

            return modules;
        }

        public static bool EnsureLoggedIn()
        {
            return EnsureLoggedIn(SuiteCRMUserSession);
        }

        public static bool EnsureLoggedIn(UserSession userSession)
        {
            bool result = false; 
            if (userSession != null)
            {
                string userId = RestAPIWrapper.GetRealUserId();
                if (string.IsNullOrEmpty(userId))
                {
                    userSession.Login();
                    result = RestAPIWrapper.GetRealUserId() != null;
                }
                else
                {
                    result = true;
                }
            }

            return result;
        }


        /// <summary>
        /// Clear the user id cache.
        /// </summary>
        public static void FlushUserIdCache()
        {
            CachedUserId = string.Empty;
        }


        /// <summary>
        /// Return the CRM id of the current user.
        /// </summary>
        /// <returns>the CRM id of the current user.</returns>
        public static string GetUserId()
        {
            if (string.IsNullOrEmpty(CachedUserId))
            {
                CachedUserId = GetRealUserId();
            }
            return CachedUserId;
        }


        /// <summary>
        /// Get the CRM id of the current user, ignoring the cache.
        /// </summary>
        /// <returns>the CRM id of the current user.</returns>
        private static string GetRealUserId()
        {
            string userId = string.Empty;

            if (SuiteCRMUserSession != null)
            {
                try
                {
                    object data = new
                    {
                        @session = SuiteCRMUserSession.id
                    };
                    userId = SuiteCRMUserSession.RestServer.GetCrmResponse<string>("get_user_id", data);
                }
                catch (Exception)
                {
                    // Swallow exception(!)
                }
            }

            return userId;
        }

        /// <summary>
        /// Sets an entry in CRM and returns the id. 'Unsafe' because if it fails (for 
        /// whatever reason), it returns the empty string. Most code which uses it fails
        /// to check for the 'empty string' return result. Use 'SetEntry' instead (which
        /// throws an exception on failure).
        /// </summary>
        /// <param name="data"></param>
        /// <param name="moduleName"></param>
        /// <returns>the CRM id of the object created or modified.</returns>
        public static string SetEntryUnsafe(NameValue[] data, string moduleName = "Emails")
        {
            try
            {
                return SetEntry(data, moduleName);
            }
            catch (System.Exception)
            {
                // Swallow exception(!)
                return string.Empty;
            }
        }

        /// <summary>
        /// Sets an entry in CRM and returns the id. 'Unsafe' because if it fails (for 
        /// whatever reason), it returns the empty string. Most code which uses it fails
        /// to check for the 'empty string' return result. Use 'SetEntry' instead (which
        /// throws an exception on failure).
        /// </summary>
        /// <param name="data"></param>
        /// <param name="moduleName"></param>
        /// <returns>the CRM id of the object created or modified.</returns>
        public static string SetEntryUnsafe(NameValueCollection data, string moduleName = "Emails")
        {
            return SetEntryUnsafe(data.ToArray(), moduleName);
        }


        public static string SetEntry(NameValue[] values, string moduleName)
        {
            EnsureLoggedIn();
            object data = new
            {
                @session = SuiteCRMUserSession.id,
                @module_name = moduleName,
                @name_value_list = values
            };
            SetEntryResult _result = SuiteCRMUserSession.RestServer.GetCrmResponse<SetEntryResult>("set_entry", data);
            return _result.id == null ?
                string.Empty :
                _result.id.ToString();
        }

        /// <summary>
        /// Send acceptance status to CRM to synchronise.
        /// </summary>
        /// <param name="meetingId">The id of the meeting to accept an invitation to.</param>
        /// <param name="moduleName">The module within which the invitee resides.</param>
        /// <param name="moduleId">The id of the invitee within that module.</param>
        /// <param name="status">The status to set.</param>
        /// <returns>true if nothing dreadful happens - not necessarily proof that the call succeeded.</returns>
        public static bool AcceptDeclineMeeting(string meetingId, string moduleName, string moduleId, string status)
        {
            if (moduleName.EndsWith("s"))
            {
                moduleName = moduleName.Substring(0, moduleName.Length - 1);
            }
            String pathPart = 
                $"index.php?entryPoint=acceptDecline&module=Meetings&{moduleName.ToLower()}_id={moduleId}&record={meetingId}&accept_status={status}";

            EnsureLoggedIn();
            return SuiteCRMUserSession.RestServer.SendGetRequest(pathPart);
        }

        public static string GetRelationship(string MainModule, string ID, string ModuleToFind)
        {
            try
            {
                EnsureLoggedIn();
                object data = new
                {
                    @session = SuiteCRMUserSession.id,
                    @module_name = MainModule,
                    @module_id = ID,
                    @link_field_name = ModuleToFind,
                    @related_module_query = "",
                    @related_fields = new string[] { "id" }
                };
                Relationships _result = SuiteCRMUserSession.RestServer.GetCrmResponse<Relationships>("get_relationships", data);
                if (_result.entry_list.Length > 0)
                    return _result.entry_list[0].id;
                return "";
            }
            catch (System.Exception)
            {
                // Swallow exception(!)
                return "";
            }
        }

        public static EntryValue[] GetRelationships(string MainModule, string ID, string ModuleToFind, string[] fields)
        {
            try
            {
                EnsureLoggedIn();
                object data = new
                {
                    @session = SuiteCRMUserSession.id,
                    @module_name = MainModule,
                    @module_id = ID,
                    @link_field_name = ModuleToFind,
                    @related_module_query = "",
                    @related_fields = fields
                };
                Relationships _result = SuiteCRMUserSession.RestServer.GetCrmResponse<Relationships>("get_relationships", data);
                
                return _result.entry_list.Length > 0 ? 
                    _result.entry_list:
                    null;
            }
            catch (System.Exception any)
            {
                Log.Error($"RestAPIWrapper.GetRelationships: main `{MainModule}`, id `{ID}`, seeking `{ModuleToFind}`.", any);
                return null;
            }
        }

        /// <summary>
        /// Sets a CRM relationship and returns boolean success. 'Unsafe' because most 
        /// callers ignore the result. Call 'SetRelationship' instead, which throws an
        /// exception on failure.
        /// </summary>
        /// <param name="relationship">The relationship to set.</param>
        public static bool SetRelationshipUnsafe(SetRelationshipParams relationship)
        {
            bool result;

            try
            {
                result = TrySetRelationship(relationship, Objective.Meeting);

                if (!result)
                {
                    Log.Warn("SuiteCrmHelper.SetRelationshipUnsafe: failed to set relationship");
                }
            }
            catch (System.Exception exception)
            {
                Log.Error("SuiteCrmHelper.SetRelationshipUnsafe:", exception);
                result = false;
            }

            return result;
        }

        /// <summary>
        /// The protocols for how link fields are named vary. Try the most likely two possibilities,
        /// and log failures.
        /// </summary>
        /// <param name="relationship">The relationship to set.</param>
        /// <returns>True if the relationship was created, else false.</returns>
        public static bool TrySetRelationship(SetRelationshipParams relationship, Objective objective)
        {
            return TrySetRelationship(relationship, $"{relationship.module2}") ||
                TrySetRelationship(relationship, $"{relationship.module2}_{relationship.module1}") ||
                TrySetRelationship(relationship, GetActivitiesLinks(relationship.module1, objective));
        }

        /// <summary>
        /// Try, in turn, each field in this list of candidate fields seeking one which allows a relationship 
        /// to be successfully created.
        /// </summary>
        /// <remarks>
        /// If the common relationship field names don't work, brute force it by getting all the possibles.
        /// </remarks>
        /// <param name="relationship">The relationship we're trying to make.</param>
        /// <param name="candidateFields">Fields through which the relationship might be made.</param>
        /// <returns>True if the relationship was made.</returns>
        private static bool TrySetRelationship(SetRelationshipParams relationship, IEnumerable<Field> candidateFields)
        {
            bool result = false;

            foreach (Field field in candidateFields)
            {
                result |= TrySetRelationship(relationship, field.name.ToLower());

                if (result) break;
            }

            return result;
        }

        /// <summary>
        /// The protocols for how link fields are named vary. Try this possibility,
        /// and log failures.
        /// </summary>
        /// <param name="relationship">The relationship to set.</param>
        /// <param name="linkFieldName">The link field name to try.</param>
        /// <returns>True if the relationship was created, else false.</returns>
        public static bool TrySetRelationship(SetRelationshipParams relationship, string linkFieldName)
        {
            bool result;

            linkFieldName = linkFieldName.ToLower();

            if (EnsureLoggedIn())
            {
                object data = new
                {
                    @session = SuiteCRMUserSession.id,
                    @module_name = relationship.module1,
                    @module_id = relationship.module1_id,
                    @link_field_name = linkFieldName,
                    @related_ids = new string[] { relationship.module2_id },
                    @name_value_list = new NameValue[] { },
                    @delete = relationship.delete
                };
                var value = SuiteCRMUserSession.RestServer.GetCrmResponse<RESTObjects.eNewSetRelationshipListResult>("set_relationship", data);

                if (value.Failed == 0)
                {
                    Log.Info($"SuiteCrmHelper.SetRelationship: successfully set relationship using link field name '{linkFieldName}'");
                }
                else
                {
                    Log.Warn($"SuiteCrmHelper.SetRelationship: failed to set relationship using link field name '{linkFieldName}'");
                }

                result = (value.Created != 0);
            }
            else
            {
                result = false;
            }

            return result;
        }


        public static NameValue SetNameValuePair(string name, object value)
        {
            return new NameValue { name = name, value = value };
        }       

        /// <summary>
        /// Perform a get_server_info call, and return the result.
        /// </summary>
        /// <returns>the result of the get_server_info call.</returns>
        public static RESTObjects.ServerInfo GetServerInfo()
        {
            object data = new
            {
                @session = SuiteCRMUserSession.id
            };

            return SuiteCRMUserSession.RestServer.GetCrmResponse<RESTObjects.ServerInfo>("get_server_info", data);
        }
        
        public static EntryList GetEntryList(string module, string query, int limit, string order_by, int offset, bool GetDeleted, string[] fields)
        {
            EnsureLoggedIn();
            EntryList result = new EntryList();
            object data = new
            {
                @session = SuiteCRMUserSession.id,
                @module_name = module,
                @query = query,
                @order_by = order_by,
                @offset = offset,
                @select_fields = fields,
                @link_names_to_fields_array = module == "Meetings" ?
                new[] {
                    new { @name = "users", @value = new[] {"id", "email1" } },
                    new { @name = "contacts", @value = new[] {"id", "account_id", "email1" } },
                    new { @name = "leads", @value = new[] {"id", "email1" } }
                } :
                null,
                @max_results = $"{limit}",
                @deleted = GetDeleted
            };
            result = SuiteCRMUserSession.RestServer.GetCrmResponse<RESTObjects.EntryList>("get_entry_list", data);                
            if (result.error != null)
            {
                throw new Exception(result.error.description);                    
            }

            if (result.entry_list != null)
            {
                try
                {
                    result.resolveLinks();
                    Hashtable hashtable = new Hashtable();
                    int index = 0;
                    foreach (EntryValue _value in result.entry_list)
                    {
                        if (!hashtable.Contains(_value.id))
                        {
                            hashtable.Add(_value.id, _value);
                        }
                        result.entry_list[index] = null;
                        index++;
                    }
                    int num2 = 0;
                    result.entry_list = null;
                    result.entry_list = new EntryValue[hashtable.Count];
                    result.result_count = hashtable.Count;
                    foreach (DictionaryEntry entry in hashtable)
                    {
                        result.entry_list[num2] = (EntryValue)entry.Value;
                        num2++;
                    }
                }
                catch (System.Exception)
                {
                    result.result_count = 0;
                }
            }

            return result;
        }

        public static string GetValueByKey(EntryValue entry, string key)
        {
            string str = string.Empty;
            foreach (NameValue _value in entry.nameValueList)
            {
                if (_value.name == key)
                {
                    str = _value.value.ToString();
                }
            }
            return str;
        }

        /// <summary>
        /// Get the module fields data for the module with this name, if any.
        /// </summary>
        /// <param name="module">the name of the module to query.</param>
        /// <returns>A structure of module's fields data.</returns>
        public static ModuleFields GetFieldsForModule(string module)
        {
            ModuleFields result;

            if (RestAPIWrapper.moduleFieldsCache.ContainsKey(module))
            {
                result = RestAPIWrapper.moduleFieldsCache[module];
            }
            else
            {
                if (!string.IsNullOrEmpty(module) && SuiteCRMUserSession != null)
                {
                    EnsureLoggedIn();
                    object data = new
                    {
                        @session = SuiteCRMUserSession.id,
                        @module_name = module
                    };

                    result = SuiteCRMUserSession.RestServer.GetCrmResponse<ModuleFields>("get_module_fields", data);
                    RestAPIWrapper.moduleFieldsCache[module] = result;
                }
                else
                {
                    result = new ModuleFields();
                }
            }

            return result;
        }

        public static List<string> GetFields(string module)
        {
            List<string> list = new List<string>();

            foreach (Field field in GetFieldsForModule(module).moduleFields)
            {
                list.Add(field.name);
            }
            return list;
        }

        /// <summary>
        /// Get the names of all the fields of the module with this name whose data type is char or varchar or name.
        /// </summary>
        /// <param name="module">The module whose field names we're interested in.</param>
        /// <returns>The names of the character fields.</returns>
        public static List<string> GetCharacterFields(string module)
        {
            List<string> list = new List<string>();

            foreach (Field field in GetFieldsForModule(module).moduleFields)
            {
                if (!field.name.EndsWith("_c"))
                {
                    /* fields with names ending '_c' are generally in a separate 'custom' table */
                    switch (field.type)
                    {
                        case "char":
                        case "email":
                        case "fullname":
                        case "name":
                        case "phone":
                        case "readonly":
                        case "text":
                        case "url":
                        case "varchar":
                            /* these are fields we can search for string data */
                            list.Add(field.name);
                            break;
                        case "assigned_user_name":
                        case "bool":
                        case "currency":
                        case "date":
                        case "datetime":
                        case "enum":
                        case "float":
                        case "id":
                        case "image": /* you could search image fields but it's 
                        * unlikely to be useful */
                        case "int":
                        case "longtext": /* probably safer not to search this */
                        case "relate":
                            /* these are not */
                            break;
                        default:
                            Log.Debug($"Unknown field type {field.type}");
                            break;
                    }
                }
            }
            return list;
        }

        /// <summary>
        /// Find the fields, among the fields of this module, which are links and where
        /// the name of the relationship linked contains the token '_activities_', prioritising
        /// those which also contain this objective.
        /// </summary>
        /// <param name="module">The name of the module to examine.</param>
        /// <param name="objective">The objective we're seeking in the relationship.</param>
        /// <returns>Its activities link fields.</returns>
        public static IEnumerable<Field> GetActivitiesLinks(string module, Objective objective)
        {
            var linkFields = GetFieldsForModule(module).linkFields;
            var objectiveName = objective.ToString().ToLower();
            IEnumerable<Field> result = GetSubstringsLinks(linkFields, new List<string>() { "_activities_", objectiveName });

            if (result.Count() == 0)
            {
                /* failed to find a relationship with both _activities_ and the objective */
                result = GetSubstringsLinks(linkFields, new List<string>() { objectiveName });
            }

            return result;
        }

        /// <summary>
        /// Filter from these link fields those whose relationship names contain all of these substrings
        /// </summary>
        /// <param name="linkFields">The link fields to filter.</param>
        /// <param name="substrings">The strings to filter them by.</param>
        /// <returns>The fields whose relationship names contain all of these substrings.</returns>
        private static IEnumerable<Field> GetSubstringsLinks(IEnumerable<Field> linkFields, IEnumerable<string> substrings)
        {
            return linkFields.Where(l => l.type.Equals("link") && StringContainsAll(l.relationship, substrings));
        }

        /// <summary>
        /// Return true if this target contains all these substrings.
        /// </summary>
        /// <param name="target">The target string.</param>
        /// <param name="substrings">The substrings.</param>
        /// <returns>true if this target contains all these substrings.</returns>
        private static bool StringContainsAll(string target, IEnumerable<string> substrings)
        {
            return string.IsNullOrEmpty(target) ?
                false :
                substrings.Where(s => target.Contains(s)).Count() == substrings.Count();
        }

        /// <remarks>
        /// TODO: This really should be a data table, not code.
        /// </remarks>
        public static string[] GetSugarFields(string module)
        {
            string[] strArray = new string[14];
            if (module == null)
            {
                return strArray;
            }
            if (module == "Contacts")
            {
                return new string[] { 
                    "id", "first_name", "last_name", "email1", "phone_work", "phone_home", "title", "department", "primary_address_city", "primary_address_country", "primary_address_postalcode", "primary_address_state", "primary_address_street", "description", "user_sync", "date_modified", 
                    "account_name", "phone_mobile", "phone_fax", "salutation", "sync_contact"
                 };
            }
            if (module == "Tasks")
            {
                return new string[] { "id", "name", "description", "date_due", "status", "date_modified", "date_start", "priority", "assigned_user_id" };
            }
            if (module == "Meetings")
            {
                return new string[] { "id", "name", "description", "date_start", "date_end", "location", "date_modified", "duration_minutes", "duration_hours", "invitees", "assigned_user_id", "outlook_id" };
            }
            if (module == "Calls")
            {
                return new string[] { "id", "name", "description", "date_start", "date_end", "date_modified", "duration_minutes", "duration_hours" };
            }
            return strArray;
        }
    }
}
