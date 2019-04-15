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

namespace SuiteCRMAddIn.Daemon
{
    using SuiteCRMAddIn.Exceptions;
    using SuiteCRMClient.Logging;
    using System;

    internal class DeferredStartupAction : AbstractDaemonAction

    {

        private readonly ILogger Log = Globals.ThisAddIn.Log;

        public DeferredStartupAction() : base(2 * Properties.Settings.Default.StartupDeferral * 1000 / DaemonWorker.Instance.IntervalMs)
        {
        }

        /// <summary> 
        /// Wait until startup deferral period has expired, and then start. 
        /// </summary> 
        /// <returns></returns> 
        public override string Perform()
        {
            string result = "DeferredStartupAction.Perform: still waiting";

            if (this.Attempts < (Properties.Settings.Default.StartupDeferral * 1000 / DaemonWorker.Instance.IntervalMs))
            {
                throw new ActionRetryableException($"{result} ({this.Attempts}).");
            }
            else
            {
                try
                {
                    Globals.ThisAddIn.DeferredStartup();

                    result = "DeferredStartupAction.Perform: complete.";
                }
                catch (Exception any)
                {
                    Log.Error("Failure in DeferredStartupAction.Perform", any);

                    throw new ActionFailedException("Failure in DeferredStartupAction.Perform", any);
                }
            }

            return result;
        }
    }
}