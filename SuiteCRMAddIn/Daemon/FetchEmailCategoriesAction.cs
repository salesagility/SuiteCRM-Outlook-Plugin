
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
    using BusinessLogic;
    using Exceptions;
    using SuiteCRMClient;
    using SuiteCRMClient.RESTObjects;
    using System.Linq;
    using System.Net;

    /// <summary>
    /// The idea here is that at startup, the addin fires off one instance of this action, 
    /// passing in its (initially empty) list of email categories by reference. When this
    /// action runs, it populates that list, so that if the email archiving dialogue is
    /// subsequently opened, it has the appropriate values.
    /// </summary>
    public class FetchEmailCategoriesAction : AbstractDaemonAction
    {
        /// <summary>
        /// Construct a new instance of the FetchEmailCategoriesAction class.
        /// </summary>
        public FetchEmailCategoriesAction() : base(5)
        {
            if (Properties.Settings.Default.EmailCategories == null)
            {
                Properties.Settings.Default.EmailCategories = new EmailCategoriesCollection();
            }
        }

        /// <summary>
        /// Replace the items in my items list, which is the list passed in by the caller, with
        /// the options returned for the 'category_id' field of the 'email' module.
        /// </summary>
        public override string Perform()
        {
            try
            {
                Field field = RestAPIWrapper.GetFieldsForModule("Emails").moduleFields.FirstOrDefault(x => x.name == "category_id");

                if (field != null)
                {
                    Properties.Settings.Default.EmailCategories.IsImplemented = true;
                    Properties.Settings.Default.EmailCategories.Clear();
                    Properties.Settings.Default.EmailCategories.AddRange(field.Options.Keys.OrderBy(x => x));
                }
                else
                {
                    /* the CRM instance does not have the category_id field in its emails module */
                    Properties.Settings.Default.EmailCategories.IsImplemented = false;
                }

                Properties.Settings.Default.Save();

                return "OK";
            } 
            catch (WebException wex)
            {
                if (wex.Status == WebExceptionStatus.Timeout)
                {
                    throw new ActionRetryableException("Temporary network error", wex);
                }
                else
                {
                    throw;
                }
            }
        }
    }
}
