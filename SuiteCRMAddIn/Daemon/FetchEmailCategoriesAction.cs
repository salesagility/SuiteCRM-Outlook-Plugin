
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
    using SuiteCRMClient;
    using SuiteCRMClient.RESTObjects;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// The idea here is that at startup, the addin fires off one instance of this action, 
    /// passing in its (initially empty) list of email categories by reference. When this
    /// action runs, it populates that list, so that if the email archiving dialogue is
    /// subsequently opened, it has the appropriate values.
    /// </summary>
    public class FetchEmailCategoriesAction : AbstractDaemonAction
    {
        /// <summary>
        /// The list of items I shall modify.
        /// </summary>
        private readonly List<string> items;

        /// <summary>
        /// Construct a new instance of the FetchEmailCategoriesAction class.
        /// </summary>
        /// <param name="listToModify">The list of items I shall modify</param>
        public FetchEmailCategoriesAction(List<string> listToModify) : base(5)
        {
            this.items = listToModify;
        }

        /// <summary>
        /// Add the options returned for the 'category' field of the 'email' module to my items
        /// list, which is the list passed in by my caller.
        /// </summary>
        public override void Perform()
        {
            //eModuleFields fields = clsSuiteCRMHelper.GetFieldsForModule("Emails");
            //eField field = fields.moduleFields.FirstOrDefault(x => x.name == "category_id");

            //items.AddRange(field.options.Keys.OrderBy(x => x));
        }
    }
}
