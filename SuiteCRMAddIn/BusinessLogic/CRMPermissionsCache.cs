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
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Cache of CRM import/export permissions, flushed hourly.
    /// </summary>
    /// <typeparam name="OutlookItemType">The type of outlook item for which I manage 
    /// permissions (may be stored in more than one module).</typeparam>
    public class CRMPermissionsCache<OutlookItemType> : RepeatingProcess
        where OutlookItemType : class
    {
        /// <summary>
        /// The token used by CRM to indicate import permissions.
        /// </summary>
        public const string ImportPermissionToken = "import";

        /// <summary>
        /// The token used by CRM to indicate export permissions.
        /// </summary>
        public const string ExportPermissionToken = "export";

        /// <summary>
        /// A cache, by module name, of whether we have import to CRM, export from CRM,
        /// of both permissions for the given module.
        /// </summary>
        private Dictionary<string, SyncDirection.Direction> cache =
            new Dictionary<string, SyncDirection.Direction>();

        /// <summary>
        /// The synchroniser on whose behalf I work.
        /// </summary>
        private Synchroniser<OutlookItemType> synchroniser;

        /// <summary>
        /// The logger I log to.
        /// </summary>
        private ILogger log;

        /// <summary>
        /// A lock for cache access.
        /// </summary>
        /// <see cref="CRMPermissionsCache{OutlookItemType}.HasAccess(string, string)"/>
        /// <see cref="CRMPermissionsCache{OutlookItemType}.PerformIteration"/> 
        private object cacheLock = new object();

        /// <summary>
        /// Construct a new instance of a permissions cache for this syncrhoniser using this log.
        /// </summary>
        /// <param name="synchroniser">The synchroniser on whose behalf I shall work.</param>
        /// <param name="log">The logger I shall log to.</param>
        public CRMPermissionsCache(Synchroniser<OutlookItemType> synchroniser, ILogger log) : 
            base($"PC Permissions cache ${synchroniser.GetType().Name}", log)
        {
            this.synchroniser = synchroniser;
            this.log = log;
            this.SyncPeriod = TimeSpan.FromHours(1);
        }


        /// <summary>
        /// Check whether my synchroniser is allowed import access for the specified CRM module.
        /// </summary>
        /// <param name="crmModule">The name of the CRM module to check.</param>
        /// <returns>true if my synchroniser is allowed import access for the specified CRM module.</returns>
        public bool HasImportAccess(string crmModule)
        {
            return this.HasAccess(crmModule, ImportPermissionToken);
        }


        /// <summary>
        /// Check whether my synchroniser is allowed export access for its default CRM module.
        /// </summary>
        /// <returns>true if my synchroniser is allowed export access for its default CRM module.</returns>
        public bool HasExportAccess()
        {
            return this.HasExportAccess(this.synchroniser.DefaultCrmModule);
        }


        /// <summary>
        /// Check whether my synchroniser is allowed export access for the specified CRM module.
        /// </summary>
        /// <param name="crmModule">The name of the CRM module to check.</param>
        /// <returns>true if my synchroniser is allowed export access for the specified CRM module.</returns>
        public bool HasExportAccess(string crmModule)
        {
            return this.HasAccess(crmModule, ExportPermissionToken);
        }


        /// <summary>
        /// Check whether my synchroniser is allowed both import and export access for its default CRM module.
        /// </summary>
        /// <returns>true if my synchroniser is allowed both import and export access for its default CRM module.</returns>
        public bool HasImportExportAccess()
        {
            return this.HasImportExportAccess(this.synchroniser.DefaultCrmModule);
        }


        /// <summary>
        /// Check whether my synchroniser is allowed both import and export access for the specified CRM module.
        /// </summary>
        /// <param name="crmModule">The name of the CRM module to check.</param>
        /// <returns>true if my synchroniser is allowed both import and export access for the specified CRM module.</returns>
        public bool HasImportExportAccess(string crmModule)
        {
            return this.HasImportAccess(crmModule) &&
                this.HasExportAccess(crmModule);
        }


        /// <summary>
        /// Check whether my synchroniser is allowed access to the specified CRM module, with the specified permission.
        /// </summary>
        /// <remarks>
        /// <para>
        /// Note that, surprisingly, although CRM will report what permissions we have, it will not 
        /// enforce them, so we have to do the honourable thing and not cheat.
        /// </para>
        /// <para>
        /// Note also that the cache is locked here, not in lower level functions. The only other
        /// place the cache is locked is in PerformIteration, where the cache is flushed.
        /// </para>
        /// </remarks>
        /// <param name="moduleName">The name of the CRM module being queried.</param>
        /// <param name="permission">The permission sought.</param>
        /// <returns>true if my synchroniser is allowed access to the specified CRM module, with the specified permission.</returns>
        /// <see cref="CRMPermissionsCache{OutlookItemType}.PerformIteration"/> 
        public bool HasAccess(string moduleName, string permission)
        {
            bool result = false;

            lock (this.cacheLock)
            {
                bool? cached = HasCachedAccess(moduleName, permission);

                if (cached != null)
                {
                    result = (bool)cached;
                    this.log.Debug($"Permissions cache hit for {moduleName}/{permission}");
                }
                else
                {
                    this.log.Debug($"Permissions cache miss for {moduleName}/{permission}");
                    try
                    {
                        AvailableModules oList = RestAPIWrapper.GetModules();
                        bool canExport = oList.items.FirstOrDefault(a => a.module_label == moduleName)
                            ?.module_acls1.FirstOrDefault(b => b.action == ExportPermissionToken)
                            ?.access ?? false;
                        bool canImport = oList.items.FirstOrDefault(a => a.module_label == moduleName)
                            ?.module_acls1.FirstOrDefault(b => b.action == ImportPermissionToken)
                            ?.access ?? false;

                        CacheAccessPermission(moduleName, ExportPermissionToken, canExport);
                        CacheAccessPermission(moduleName, ImportPermissionToken, canImport);

                        Log.Debug($"Cached {this.cache[moduleName]} permission for {moduleName}");

                        cached = HasCachedAccess(moduleName, permission);

                        if (cached == null)
                        {
                            /* really shouldn't happen - we've just set it! */
                            Log.Warn($"Cannot detect access {moduleName}/{permission} despite having just set it");
                            /* not really satisfactory, but unlikely to happen */
                            result = false;
                        }
                        else
                        {
                            result = (bool)cached;
                        }
                    }
                    catch (Exception fetchFailed)
                    {
                        Log.Error($"Cannot detect access {moduleName}/{permission} because {fetchFailed.Message}", fetchFailed);
                    }
                }
            }

            return result;
        }


        /// <summary>
        /// Cache an access permission received from CRM, so we don't have to repeatedly request it.
        /// </summary>
        /// <remarks>
        /// The cache is modified additively. It we already know we have one permission, and find we have the
        /// other, then we assume both. There isn't presently any mechanism to remove permissions from the cache. 
        /// </remarks>
        /// <param name="moduleName">The module to which access may be granted.</param>
        /// <param name="direction">The direction in which access may be granted.</param>
        /// <param name="allowed">The access that should be granted.</param>
        private void CacheAccessPermission(string moduleName, string direction, bool allowed)
        {
            if (this.cache.ContainsKey(moduleName))
            {
                switch (this.cache[moduleName])
                {
                    case SyncDirection.Direction.Neither:
                        /* shouldn't happen as it would be unwise to cache 'neither' unless 
                         * we know it is true - which we won't. */
                        CacheAllowAccess(moduleName, direction, allowed);
                        break;
                    case SyncDirection.Direction.Export:
                        if (direction == ImportPermissionToken && allowed)
                        {
                            /* if we already had export permission and now we have import permission, we have
                             * both. */
                            this.cache[moduleName] = SyncDirection.Direction.BiDirectional;
                        }
                        break;
                    case SyncDirection.Direction.Import:
                        if (direction == ExportPermissionToken && allowed)
                        {
                            /* if we already had import permission and now we have export permission, we have
                             * both. */
                            this.cache[moduleName] = SyncDirection.Direction.BiDirectional;
                        }
                        break;
                }
            }
            else
            {
                CacheAllowAccess(moduleName, direction, allowed);
            }
        }


        /// <summary>
        /// Cache allowed access in the specified direction.
        /// </summary>
        /// <remarks>
        /// Assumes there is currently no cached value for the specified module; if there is,
        /// it will be overwritten.
        /// </remarks>
        /// <param name="moduleName">The module to which access may be granted.</param>
        /// <param name="direction">The direction in which access may be granted.</param>
        /// <param name="allowed">The access that should be granted.</param>
        private void CacheAllowAccess(string moduleName, string direction, bool allowed)
        {
            if (allowed)
            {
                this.cache[moduleName] = direction == ImportPermissionToken ?
                    SyncDirection.Direction.Import : SyncDirection.Direction.Export;
            }
        }


        /// <summary>
        /// Does the currently cached value allow access to this module name in this direction?
        /// </summary>
        /// <param name="moduleName"></param>
        /// <param name="direction"></param>
        /// <returns>True if access is permitted, false if it's denied, null if there's no 
        /// cached value.</returns>
        private bool? HasCachedAccess(string moduleName, string direction)
        {
            bool? result = null;

            if (this.cache.ContainsKey(moduleName))
            {
                SyncDirection.Direction cachedValue = this.cache[moduleName];
                result = (direction == ImportPermissionToken && SyncDirection.AllowOutbound(cachedValue)) ||
                    (direction == ExportPermissionToken && SyncDirection.AllowInbound(cachedValue));
            }

            return result;
        }


        /// <summary>
        /// Periodically flush the cache.
        /// </summary>
        /// <remarks>
        /// Note that the cache is also locked in HasAccess.
        /// </remarks>
        /// <see cref="CRMPermissionsCache{OutlookItemType}.HasAccess(string, string)"/> 
        internal override void PerformIteration()
        {
            lock (this.cacheLock)
            {
                Log.Info("Flushing permissions cache");
                cache = new Dictionary<string, SyncDirection.Direction>();
            }
        }
    }
}
