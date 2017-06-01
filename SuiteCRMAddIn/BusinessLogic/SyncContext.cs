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
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// A bundle of handles onto all the things which are needed to allow synchronisation to take place.
    /// </summary>
    public class SyncContext
    {
        private readonly Outlook.Application application;
        private Outlook.OlItemType currentFolderItemType;

        public SyncContext(Outlook.Application application)
        {
            this.application = application;
            currentFolderItemType = Outlook.OlItemType.olMailItem;
        }

        public Outlook.Application Application => application;

        public ILogger Log => Globals.ThisAddIn.Log;

        public Outlook.OlItemType CurrentFolderItemType => currentFolderItemType;

        public void SetCurrentFolder(Outlook.MAPIFolder folder)
        {
            currentFolderItemType = folder.DefaultItemType;
        }
    }
}
