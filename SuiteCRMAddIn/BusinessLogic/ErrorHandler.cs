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
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.Exceptions;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using SuiteCRMAddIn.Daemon;
    using SuiteCRMAddIn.Exceptions;


    public class ErrorHandler
    {
        private static HashSet<string> SeenExceptions = new HashSet<string>();
        public static void Handle(Exception error)
        {
            ErrorHandler.Handle("SuiteCRM Addin has encountered a problem", error);
        }

        /// <summary>
        /// Handle bad credentials specially.
        /// </summary>
        /// <param name="badCredentials"></param>
        public static void Handle(BadCredentialsException badCredentials)
        {
            if (Globals.ThisAddIn.ShowReconfigureOrDisable("Login failed; have your credentials changed?") == DialogResult.Cancel)
            {
                Globals.ThisAddIn.Disable();
            }
        }

        public static void Handle(string message)
        {
            Handle(message, (Exception)null);
        }

        public static void Handle(OutOfMemoryException error)
        {
            Handle( "SuiteSRM AddIn recovered from an out of memory error; no work was lost, but some tasks may not have been completed", error);
        }

        public static void Handle(string contextMessage, OutOfMemoryException error, bool notify = false)
        {
            Globals.ThisAddIn.Log.Error(contextMessage, error);
            MessageBox.Show(ComposeErrorDescription(contextMessage, error), "SuiteCRM  AddIn ran out of memory", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static void Handle(string contextMessage, NeverShowUserException error, bool notify = false)
        {
            Globals.ThisAddIn.Log.Error(contextMessage, error);
        }

        /// <summary>
        /// Handle this error in the context described in this contextMessage.
        /// </summary>
        /// <param name="contextMessage">A message describing what was being attempted when the error occurred.</param>
        /// <param name="error">The error.</param>
        /// <param name="notify">If true, notify the user anyway, overriding the ShowExceptions setting.</param>
        public static void Handle(string contextMessage, Exception error, bool notify = false)
        {
            Globals.ThisAddIn.Log.Error(contextMessage, error);

            if (notify)
            {
                MessageBox.Show(ComposeErrorDescription(contextMessage, error), "SuiteCRM Addin Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                switch (Properties.Settings.Default.ShowExceptions)
                {
                    case PopupWhen.Never:
                        break;
                    case PopupWhen.FirstTime:
                        var errorClassName = error?.GetType().Name ?? string.Empty;

                        if (!SeenExceptions.Contains(errorClassName))
                        {
                            SeenExceptions.Add(errorClassName);
                                MessageBox.Show(ComposeErrorDescription(contextMessage, error), "SuiteCRM Addin Error",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        break;
                    default:
                        MessageBox.Show(ComposeErrorDescription(contextMessage, error), "SuiteCRM Addin Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                }
            }
        }

        private static string ComposeErrorDescription(string contextMessage, Exception error)
        {
            StringBuilder bob = new StringBuilder(contextMessage);

            for (var e = error; e != null; e = e.InnerException)
            {
                bob.Append( e == error ? " " : " Caused by: ").Append(e.GetType().Name).Append(" '").Append(e.Message).Append("'");
            }

            return bob.ToString();
        }

        /// <summary>
        /// Do this action and, if an error occurs, invoke the error handler on it with this message.
        /// </summary>
        /// <remarks>
        /// \todo this method is duplicated in Robustness, but the copy in ErrorHandler is preferred;
        /// in the next release it is intended to remove the Robustness class and move its functionality
        /// to ErrorHandler.
        /// </remarks>
        /// <param name="action">The action to perform</param>
        /// <param name="message">A string describing what the action was intended to achieve.</param>
        public static void DoOrHandleError(Action action, string message)
        {
            try
            {
                action();
            }
            catch (Exception problem)
            {
                ErrorHandler.Handle(message, problem);
            }
        }


        public enum PopupWhen
        {
            Never,
            FirstTime,
            EveryTime
        }
    }
}
