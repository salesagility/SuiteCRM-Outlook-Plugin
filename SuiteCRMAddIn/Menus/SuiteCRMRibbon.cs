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
using Microsoft.Office.Core;
using stdole;
using SuiteCRMAddIn.Properties;
using SuiteCRMClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new SuiteCRMRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace SuiteCRMAddIn
{
    [ComVisible(true)]
    public class SuiteCRMRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        
        public SuiteCRMRibbon()
        {
        }

        public bool SuiteCRMMainTab_GetVisible(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer() != null)
            {
                if (Globals.ThisAddIn.Application.ActiveExplorer().Selection is Outlook.MailItem)
                {
                    return true;
                }
            }
            return false;
        }

        public IPictureDisp GetImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "btnArchive":
                case "btnArchive1":
                case "btnArchive2":
                case "btnAddressBook":
                    return RibbonImageHelper.Convert(Resources.SuiteCRM1);

                case "btnSettings":
                    return RibbonImageHelper.Convert(Resources.Settings);
                
            }
            return null;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string result;

            switch (ribbonID)
            {
                case "Microsoft.Outlook.Mail.Read":
                case "Microsoft.Outlook.Explorer":
                    result = (Globals.ThisAddIn.OutlookVersion <= OutlookMajorVersion.Outlook2007) ?
                        GetResourceText("SuiteCRMAddIn.Menus.MailRead2007.xml") :
                        GetResourceText("SuiteCRMAddIn.Menus.MailRead.xml");
                    break;
                case "Microsoft.Outlook.Mail.Compose":
                    result = GetResourceText("SuiteCRMAddIn.Menus.MailCompose.xml");
                    break;
                default:
                    result = String.Empty;
                    break;
            }

            return result;
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            Globals.ThisAddIn.RibbonUI = ribbonUI;
        }

        #endregion

        #region Click Events
        public void btnArchive_Action(IRibbonControl control)
        {
            DoOrLogError(() =>
                ManualArchive());
        }

        public void btnSettings_Action(IRibbonControl control)
        {
            DoOrLogError(() =>
                Globals.ThisAddIn.ShowSettingsForm());
        }

        public void btnAddressBook_Action(IRibbonControl control)
        {
            frmAddressBook objAddressBook = new frmAddressBook();
            objAddressBook.Show();
        }
      
        
        private void ManualArchive()
        {
            if (Globals.ThisAddIn.SuiteCRMUserSession.NotLoggedIn)
            {
                Globals.ThisAddIn.ShowSettingsForm();
            }
            Globals.ThisAddIn.ShowArchiveForm();
        }
        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        /// <summary>
        /// Wrapper around invocation of an action, to provide consistent logging of 
        /// otherwise-uncaught exceptions.
        /// </summary>
        /// <param name="action">The actual action handler to invoke.</param>
        private void DoOrLogError(Action action)
        {
            Robustness.DoOrLogError(Globals.ThisAddIn.Log, action);
        }
    }
}
