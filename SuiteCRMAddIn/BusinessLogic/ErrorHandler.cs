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
            Handle(message, null);
        }

        public static void Handle(string contextMessage, Exception error)
        {
            Globals.ThisAddIn.Log.Error(contextMessage, error);
            var errorClassName = error?.GetType().Name ?? string.Empty;
            StringBuilder bob = new StringBuilder(contextMessage);

            for (Exception e = error; e != null; e = e.InnerException)
            {
                if (e != error)
                {
                    bob.Append("Caused by: ");
                }
                bob.Append(e.GetType().Name).Append(e.Message);
            }
            string text = bob.ToString();

            switch (Properties.Settings.Default.ShowExceptions)
            {
                case PopupWhen.Never:
                    break;
                case PopupWhen.FirstTime:
                    if (!SeenExceptions.Contains(errorClassName))
                    {
                        SeenExceptions.Add(errorClassName);
                        MessageBox.Show(text, "SuiteCRM Addin Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;
                default:
                    MessageBox.Show(text, "SuiteCRM Addin Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
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
