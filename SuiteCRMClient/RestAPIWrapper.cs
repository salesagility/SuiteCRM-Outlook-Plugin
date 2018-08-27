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
using System;
using System.Collections.Generic;

namespace SuiteCRMClient
{
    using RESTObjects;
    using System.Collections;
    using System.Linq;
    using Logging;
    using System.Net.Mail;
    using System.Text.RegularExpressions;

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
            if (userSession != null && userSession.AwaitingAuthentication == false)
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
        /// Get the user id of the user with this email address, if any.
        /// </summary>
        /// <param name="mailAddress">the email address to seek.</param>
        /// <returns>An id if available, else the empty string.</returns>
        public static string GetUserId(MailAddress mailAddress)
        {
            string result = string.Empty;
            EntryList list = GetEntryList(
                "Users", 
                $"(users.id in (select eabr.bean_id from email_addr_bean_rel eabr INNER JOIN email_addresses ea on eabr.email_address_id = ea.id where eabr.bean_module = 'Users' and ea.email_address LIKE '%{MySqlEscape(mailAddress.ToString())}%'))", 
                0, 
                "id DESC", 
                0, 
                false, 
                new string[] { "id" });
            
            if (list.entry_list != null && list.entry_list.Any())
            {
                result = list.entry_list[0].id;
            }

            return result;
        }


        /// <summary>
        /// Get the user id of the user with this email address, if any.
        /// </summary>
        /// <param name="username">the username to seek.</param>
        /// <returns>An id if available, else the empty string.</returns>
        public static string GetUserId(string username)
        {
            string result = string.Empty;
           
            EntryList list = GetEntryList("Users", $"users.user_name LIKE '%{MySqlEscape(username)}%'", 0, "id DESC", 0, false, new string[] { "id" });

            if (list?.entry_list != null && list.entry_list.Count() > 0)
            {
                result = list.entry_list[0].id;
            }

            return result;
        }


        /// <summary>
        /// Create and return a copy of this string which escapes all characters which
        /// might render MySQL vulnerable to SQL injection attacks.
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <returns>A suitably escaped copy of the input.</returns>
        public static string MySqlEscape(string input)
        {
            return string.IsNullOrEmpty(input) ? null : Regex.Replace(input, "[\\r\\n\\x00\\x1a\\\\'\"]", @"\$0");
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
                    userId = SuiteCRMUserSession.RestServer.GetCrmStringResponse("get_user_id", data);
                }
                catch (Exception fail)
                {
                    Log.Error("", fail);
                }
            }

            return userId;
        }

        /// <summary>
        /// Sets an entry in CRM and returns the id. 
        /// </summary>
        /// <param name="data">The data to set.</param>
        /// <param name="moduleName">The name of the CRM module into which to insert it.</param>
        /// <returns>the CRM id of the object created or modified.</returns>
        public static string SetEntry(NameValueCollection data, string moduleName)
        {
            return SetEntry(data.ToArray(), moduleName);
        }

        /// <summary>
        /// Sets an entry in CRM and returns the id. 
        /// </summary>
        /// <param name="data">The data to set.</param>
        /// <param name="moduleName">The name of the CRM module into which to insert it.</param>
        /// <returns>the CRM id of the object created or modified.</returns>
        public static string SetEntry(NameValue[] values, string moduleName)
        {
            if (values == null || values.Count() == 0)
            {
                throw new MissingValuesException($"Missing values when storing an instance of '{moduleName}'");
            }
            EnsureLoggedIn();
            object data = new
            {
                @session = SuiteCRMUserSession.id,
                @module_name = moduleName,
                @name_value_list = values
            };
            SetEntryResult result = SuiteCRMUserSession.RestServer.GetCrmResponse<SetEntryResult>("set_entry", data);
            return string.IsNullOrEmpty(result.id) ?
                string.Empty :
                result.id;
        }

        /// <summary>
        /// Send acceptance status to CRM to synchronise.
        /// </summary>
        /// <param name="meetingId">The id of the meeting to accept an invitation to.</param>
        /// <param name="moduleName">The module within which the invitee resides.</param>
        /// <param name="moduleId">The id of the invitee within that module.</param>
        /// <param name="status">The acceptance status to set.</param>
        /// <returns>true if nothing dreadful happens - not necessarily proof that the call succeeded.</returns>
        public static bool SetMeetingAcceptance(string meetingId, string moduleName, string moduleId, string status)
        {
            Log.Debug($"RestApiWrapper.SetMeetingAcceptance: meetingId=`{meetingId}`; moduleName=`{moduleName}`; moduleId=`{moduleId}`; status=`{status}`");
            bool result = false;

            if (EnsureLoggedIn())
            {
                object data = new
                {
                    @session = SuiteCRMUserSession.id,
                    @module_name = "Meetings",
                    @module_id = meetingId,
                    @link_field_name = moduleName.ToLower(),
                    @related_ids = new string[] { moduleId },
                    @name_value_list = new NameValue[] { new NameValue() { name = "accept_status", value = status } },
                };
                var value = SuiteCRMUserSession.RestServer.GetCrmResponse<RESTObjects.eNewSetRelationshipListResult>("set_relationship", data);

                result = value.Failed == 0;
            }
            string success = result ? "succeeded" : "failed";
            Log.Debug($"RestApiWrapper.SetMeetingAcceptance: {success}");

            return result;
        }

        public static string GetRelationship(string MainModule, string ID, string ModuleToFind)
        {
            string result;

            try
            {
                EntryValue[] entries = RestAPIWrapper.GetRelationships(MainModule, ID, ModuleToFind, new string[] { "id" });
                result = entries.Length > 0 ? entries[0].id : string.Empty;
            }
            catch (System.Exception)
            {
                // Swallow exception(!)
                result = string.Empty;
            }

            return result;
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
                    @related_fields = fields,
                    @related_module_link_name_to_fields_array = new object[] { },
                    @deleted = false,
                    @order_by = "",
                    @offset = 0,
                    @limit = false
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
                    Log.Warn("RestAPIWrapper.SetRelationshipUnsafe: failed to set relationship");
                }
            }
            catch (System.Exception exception)
            {
                Log.Error("RestAPIWrapper.SetRelationshipUnsafe:", exception);
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
                    Log.Info($"RestAPIWrapper.TrySetRelationship: successfully set relationship using link field name '{linkFieldName}'");
                }
                else
                {
                    Log.Warn($"RestAPIWrapper.TrySetRelationship: failed to set relationship using link field name '{linkFieldName}'");
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


        /// <summary>
        /// Get the specified entry from the specified module.
        /// </summary>
        /// <param name="module">The module to be queried.</param>
        /// <param name="id">The id of the entry to return.</param>
        /// <param name="fields">The fields to return.</param>
        /// <param name="linkNamesToFieldsArray">A link object to return associated records in other modules.</param>
        /// <returns>A list of entries in the module matching the query.</returns>
        public static Entry GetEntry(string module, string id, string[] fields, object linkNamesToFieldsArray = null)
        {
            Entry result = new Entry();

            if (EnsureLoggedIn())
            {
                object data = new
                {
                    @session = SuiteCRMUserSession.id,
                    @module_name = module,
                    @id = id,
                    @select_fields = fields,
                    @link_names_to_fields_array = linkNamesToFieldsArray,
                    @track_view = false
                };
                result = SuiteCRMUserSession.RestServer.GetCrmResponse<RESTObjects.Entry>("get_entry", data);
            }

            return result;
        }

        /// <summary>
        /// Get the specified entries from the specified module.
        /// </summary>
        /// <param name="module">The module to be queried.</param>
        /// <param name="query">The query to filter by.</param>
        /// <param name="limit">The limit to the number of fields to return in a page.</param>
        /// <param name="order_by">The field(s) to order by.</param>
        /// <param name="offset">The offset of the start of the page to return in the result set.</param>
        /// <param name="getDeleted">If true, include deleted records in the result.</param>
        /// <param name="fields">The fields to return.</param>
        /// <param name="linkNamesToFieldsArray">A link object to return associated records in other modules.</param>
        /// <returns>A list of entries in the module matching the query.</returns>
        public static EntryList GetEntryList(string module, string query, int limit, string order_by, int offset, bool getDeleted, string[] fields, object linkNamesToFieldsArray = null)
        {
            EntryList result = new EntryList();

            if (EnsureLoggedIn())
            {
                object data = new
                {
                    @session = SuiteCRMUserSession.id,
                    @module_name = module,
                    @query = query,
                    @order_by = order_by,
                    @offset = offset,
                    @select_fields = fields,
                    @link_names_to_fields_array = linkNamesToFieldsArray,
                    @max_results = $"{limit}",
                    @deleted = getDeleted,
                    @favorites = false
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
                        result.ResolveLinks();
                        List<EntryValue> deduped = new List<EntryValue>(result.entry_list);
                        deduped = deduped.OrderBy(x => x.id).GroupBy(x => x.id).Select(g => g.First()).ToList();

                        if (deduped.Count() < result.entry_list.Count())
                        {
                            result.entry_list = deduped.ToArray();
                            result.result_count = result.entry_list.Count();
                        }
                    }
                    catch (System.Exception)
                    {
                        result.result_count = 0;
                    }
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
            /* module name is case sensitive, initial capital */
            string moduleName = char.ToUpper(module[0]) + module.Substring(1);

            if (RestAPIWrapper.moduleFieldsCache.ContainsKey(moduleName))
            {
                result = RestAPIWrapper.moduleFieldsCache[moduleName];
            }
            else
            {
                if (!string.IsNullOrEmpty(module) && SuiteCRMUserSession != null)
                {
                    EnsureLoggedIn();
                    object data = new
                    {
                        @session = SuiteCRMUserSession.id,
                        @module_name = moduleName
                    };

                    result = SuiteCRMUserSession.RestServer.GetCrmResponse<ModuleFields>("get_module_fields", data);
                    RestAPIWrapper.moduleFieldsCache[moduleName] = result;
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
            string[] result = new string[14];

            switch (module)
            {
                case "Calls":
                    result = new string[] { "id", "name", "description", "date_start", "date_end",
                        "date_modified", "duration_minutes", "duration_hours" };
                    break;
                case "Contacts":
                    result = new string[] {"id", "first_name", "last_name", "email1", "phone_work",
                        "phone_home", "title", "department", "primary_address_city", "primary_address_country",
                        "primary_address_postalcode", "primary_address_state", "primary_address_street",
                        "description", "user_sync", "date_modified", "account_name", "phone_mobile",
                        "phone_fax", "salutation", "sync_contact" };
                    break;
                case "Meetings":
                    result = new string[] { "id", "name", "description", "date_start", "date_end", "location",
                        "date_modified", "duration_minutes", "duration_hours", "invitees", "assigned_user_id",
                        "outlook_id" };
                    break;
                case "Tasks":
                    result = new string[] { "id", "name", "description", "date_due", "status", "date_modified",
                        "date_start", "priority", "assigned_user_id" };
                    break;
                default:
                    result = new string[14];
                    break;
            }

            return result;
        }
    }
}
