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
    /// <remarks>
    /// <para>
    /// The cache itself is global, and is held on a class (static) variable; all methods of instances of 
    /// this class (actually subclass instances, since the class itself is abstract) access that single 
    /// static cache.
    /// </para>
    /// <para>
    /// NOTE that there is a significant difference between <code>module_name</code> and <code>module_key</code>.
    /// <code>module_key</code> is the name used internally by CRM for the module, and is unaffected by natural
    /// language changes; <code>module_name</code> will change with the natural language of the installation.
    /// Thanks to Andreas Ravnestad for this insight.
    /// </para>
    /// </remarks>
    public abstract class CRMPermissionsCache : RepeatingProcess
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
        /// <remarks>
        /// Since the cache holds information for all modules, it makes no sense to have 
        /// a separate cache for each instance of this class.
        /// </remarks>
        private static Dictionary<string, SyncDirection.Direction> cache =
            new Dictionary<string, SyncDirection.Direction>();


        /// <summary>
        /// The logger I log to.
        /// </summary>
        private ILogger log;

        /// <summary>
        /// A lock for cache access.
        /// </summary>
        /// <see cref="CRMPermissionsCache{OutlookItemType}.HasAccess(string, string)"/>
        /// <see cref="CRMPermissionsCache{OutlookItemType}.PerformIteration"/> 
        private static object cacheLock = new object();


        /// <summary>
        /// Construct a new instance of a permissions cache with this name and using this log.
        /// </summary>
        /// <param name="name">The name to log.</param>
        /// <param name="log">The logger I shall log to.</param>
        public CRMPermissionsCache(string name, ILogger log) : base(name, log)
        {
            this.log = log;
            this.Interval = TimeSpan.FromHours(1);
        }


        /// <summary>
        /// Check whether my synchroniser is allowed import access for the specified CRM module.
        /// </summary>
        /// <param name="moduleKey">The key of the CRM module to check.</param>
        /// <returns>true if my synchroniser is allowed import access for the specified CRM module.</returns>
        public bool HasImportAccess(string moduleKey)
        {
            return this.HasAccess(moduleKey, ImportPermissionToken);
        }


        /// <summary>
        /// Check whether my synchroniser is allowed export access for the specified CRM module.
        /// </summary>
        /// <param name="moduleKey">The key of the CRM module to check.</param>
        /// <returns>true if my synchroniser is allowed export access for the specified CRM module.</returns>
        public bool HasExportAccess(string moduleKey)
        {
            return this.HasAccess(moduleKey, ExportPermissionToken);
        }


        /// <summary>
        /// Check whether my synchroniser is allowed both import and export access for the specified CRM module.
        /// </summary>
        /// <param name="moduleKey">The key of the CRM module to check.</param>
        /// <returns>true if my synchroniser is allowed both import and export access for the specified CRM module.</returns>
        public bool HasImportExportAccess(string moduleKey)
        {
            return this.HasImportAccess(moduleKey) &&
                this.HasExportAccess(moduleKey);
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
        /// <param name="moduleKey">The key of the CRM module being queried.</param>
        /// <param name="permission">The permission sought.</param>
        /// <returns>true if my synchroniser is allowed access to the specified CRM module, with the specified permission.</returns>
        /// <see cref="CRMPermissionsCache{OutlookItemType}.PerformIteration"/> 
        public bool HasAccess(string moduleKey, string permission)
        {
            bool result = false;

            try
            {
                lock (CRMPermissionsCache.cacheLock)
                {
                    bool? cached = HasCachedAccess(moduleKey, permission);

                    if (cached != null)
                    {
                        result = (bool)cached;
                        this.Log.Debug($"Permissions cache hit for {moduleKey}/{permission}");
                    }
                    else
                    {
                        this.Log.Debug($"Permissions cache miss for {moduleKey}/{permission}");
                        result = CacheAllAndCheckAccess(moduleKey, permission);
                    }
                }

                return result;
            }
            catch (KeyNotFoundException knf)
            {
                // OK, this is impossible. But it IS happening... why?
                ErrorHandler.Handle($"Failed while checking permission {permission} for module '{moduleKey}'", knf);
                return result;
            }
        }


        /// <summary>
        /// Cache permissions for all modules, and return the value for the specified parameters.
        /// </summary>
        /// <remarks>
        /// We deliberately cache permissions for all named modules whether we're interested in them or not - 
        /// it's quicker than filtering them.
        /// </remarks>
        /// <param name="moduleKey">The key of the module we're actually seeking.</param>
        /// <param name="permission">The permission we're actually seeking.</param>
        /// <returns>true if we have the access specified by permission for the module specified by this module key.</returns>
        private bool CacheAllAndCheckAccess(string moduleKey, string permission)
        {
            bool? cached;
            bool result;

            try
            {
                this.Log.Debug("Note: ");

                foreach (AvailableModule item in RestAPIWrapper.GetModules().items)
                {
                    CachePermissionsForModule(item);
                }

                cached = HasCachedAccess(moduleKey, permission);

                if (cached == null)
                {
                    /* really shouldn't happen - we've just set it! */
                    Log.Warn($"Cannot detect access {moduleKey}/{permission} despite having just set it");
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
                ErrorHandler.Handle($"Failed to detect access {moduleKey}/{permission}", fetchFailed);
                throw;
            }

            return result;
        }


        /// <summary>
        /// Cache both inport and export permissions for the specified module.
        /// </summary>
        /// <param name="module">The modules for which permissions should be cached.</param>
        private void CachePermissionsForModule(AvailableModule module)
        {
            if (!string.IsNullOrWhiteSpace(module.module_key))
            {
                CacheAccessPermission(
                    module.module_key,
                    ImportPermissionToken,
                    module.module_acls1.FirstOrDefault(b => b.action == ImportPermissionToken)?.access ?? false);
                CacheAccessPermission(
                    module.module_key,
                    ExportPermissionToken,
                    module.module_acls1.FirstOrDefault(b => b.action == ExportPermissionToken)?.access ?? false);

                try
                {
                    Log.Debug($"Cached {CRMPermissionsCache.cache[module.module_key]} permission for {module.module_key}");
                }
                catch (KeyNotFoundException)
                {
                    // shouldn't ever happen. Ignore for now.
                }
            }
        }


        /// <summary>
        /// Cache an access permission received from CRM, so we don't have to repeatedly request it.
        /// </summary>
        /// <remarks>
        /// <para>The cache is modified additively. It we already know we have one permission, and find we have the
        /// other, then we assume both. There isn't presently any mechanism to remove permissions from the cache. </para>
        /// <para>Should never throw a KeyNotFoundException.</para>
        /// </remarks>
        /// <param name="moduleKey">The module to which access may be granted.</param>
        /// <param name="direction">The direction in which access may be granted.</param>
        /// <param name="allowed">The access that should be granted.</param>
        private void CacheAccessPermission(string moduleKey, string direction, bool allowed)
        {
            if (CRMPermissionsCache.cache.ContainsKey(moduleKey))
            {
                switch (CRMPermissionsCache.cache[moduleKey])
                {
                    case SyncDirection.Direction.Neither:
                        /* shouldn't happen as it would be unwise to cache 'neither' unless 
                         * we know it is true - which we won't. */
                        CacheAllowAccess(moduleKey, direction, allowed);
                        break;
                    case SyncDirection.Direction.Export:
                        if (direction == ImportPermissionToken && allowed)
                        {
                            /* if we already had export permission and now we have import permission, we have
                             * both. */
                            CRMPermissionsCache.cache[moduleKey] = SyncDirection.Direction.BiDirectional;
                        }
                        break;
                    case SyncDirection.Direction.Import:
                        if (direction == ExportPermissionToken && allowed)
                        {
                            /* if we already had import permission and now we have export permission, we have
                             * both. */
                            CRMPermissionsCache.cache[moduleKey] = SyncDirection.Direction.BiDirectional;
                        }
                        break;
                }
            }
            else
            {
                CacheAllowAccess(moduleKey, direction, allowed);
            }
        }


        /// <summary>
        /// Cache allowed access in the specified direction.
        /// </summary>
        /// <remarks>
        /// Assumes there is currently no cached value for the specified module; if there is,
        /// it will be overwritten.
        /// </remarks>
        /// <param name="moduleKey">The module to which access may be granted.</param>
        /// <param name="direction">The direction in which access may be granted.</param>
        /// <param name="allowed">The access that should be granted.</param>
        private void CacheAllowAccess(string moduleKey, string direction, bool allowed)
        {
            if (allowed)
            {
                CRMPermissionsCache.cache[moduleKey] = direction == ImportPermissionToken ?
                    SyncDirection.Direction.Import : SyncDirection.Direction.Export;
            }
        }


        /// <summary>
        /// Does the currently cached value allow access to this module name in this direction?
        /// </summary>
        /// <remarks>
        /// Should never throw a KeyNotFoundException.
        /// </remarks>
        /// <param name="moduleKey">The module to which access may be granted.</param>
        /// <param name="direction">The direction in which access may be granted.</param>
        /// <returns>True if access is permitted, false if it's denied, null if there's no 
        /// cached value.</returns>
        private bool? HasCachedAccess(string moduleKey, string direction)
        {
            bool? result = null;

            try
            {
                if (CRMPermissionsCache.cache.ContainsKey(moduleKey))
                {
                    SyncDirection.Direction cachedValue = CRMPermissionsCache.cache[moduleKey];
                    result = (direction == ImportPermissionToken && SyncDirection.AllowOutbound(cachedValue)) ||
                        (direction == ExportPermissionToken && SyncDirection.AllowInbound(cachedValue));
                }
            }
            catch (Exception any)
            {
                ErrorHandler.Handle($"Failed while checking the permissions cache for access to {moduleKey}/{direction}", any);
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
            Log.Info("Flushing permissions cache");
            lock (CRMPermissionsCache.cacheLock)
            {
                // no point flushing a cache that's already empty (although equally it would do little harm).
                if (cache.Keys.Count > 0)
                {
                    cache = new Dictionary<string, SyncDirection.Direction>();
                }
            }
        }
    }

    /// <summary>
    /// A thin wrapper around CRMPermissionsCache, handling permissions for a specific
    /// Outlook item type.
    /// </summary>
    /// <typeparam name="OutlookItemType">The type of outlook item for which I manage 
    /// permissions (may be stored in more than one module).</typeparam>
    /// <typeparam name="SyncStateType">The appropriate SyncState type for that Outlook item type.</typeparam>
    public class CRMPermissionsCache<OutlookItemType, SyncStateType> : CRMPermissionsCache
        where OutlookItemType : class
        where SyncStateType : SyncState<OutlookItemType>
    {
        /// <summary>
        /// The synchroniser on whose behalf I work.
        /// </summary>
        private Synchroniser<OutlookItemType, SyncStateType> synchroniser;

        /// <summary>
        /// Construct a new instance of a permissions cache for this syncrhoniser using this log.
        /// </summary>
        /// <param name="synchroniser">The synchroniser on whose behalf I shall work.</param>
        /// <param name="log">The logger I shall log to.</param>
        public CRMPermissionsCache(Synchroniser<OutlookItemType,SyncStateType> synchroniser, ILogger log) : 
            base($"PC Permissions cache ${synchroniser.GetType().Name}", log)
        {
            this.synchroniser = synchroniser;
        }


        /// <summary>
        /// Check whether my synchroniser is allowed import access for its default CRM module.
        /// </summary>
        /// <returns>true if my synchroniser is allowed import access for its default CRM module.</returns>
        public bool HasImportAccess()
        {
            return this.HasImportAccess(this.synchroniser.DefaultCrmModule);
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
        /// Check whether my synchroniser is allowed both import and export access for its default CRM module.
        /// </summary>
        /// <returns>true if my synchroniser is allowed both import and export access for its default CRM module.</returns>
        public bool HasImportExportAccess()
        {
            return this.HasImportExportAccess(this.synchroniser.DefaultCrmModule);
        }
    }
}
