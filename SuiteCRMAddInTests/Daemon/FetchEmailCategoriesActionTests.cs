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
namespace SuiteCRMAddIn.Daemon.Tests
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using SuiteCRMAddIn.Daemon;
    using SuiteCRMAddIn.Tests;

    /// <summary>
    /// Test that FetchEmailCategoriesAction actually fetches some categories.
    /// </summary>
    [TestClass()]
    public class FetchEmailCategoriesActionTests : AbstractWithCrmConnectionTest
    {
        /// <summary>
        /// The action I test.
        /// </summary>
        public FetchEmailCategoriesAction action { get; private set; }

        /// <summary>
        /// The settings which performing my action should modify.
        /// </summary>
//        private readonly clsSettings settings = new clsSettings();

        /// <summary>
        /// Specialisation: I need an action.
        /// </summary>
        [TestInitialize()]
        public override void Initialize()
        {
            base.Initialize();
//            this.action = new FetchEmailCategoriesAction(settings);
        }

        /// <summary>
        /// After performing my action, there should be some categories in my settings.
        /// </summary>
        [TestMethod()]
        public void FetchEmailCategoriesActionPerformTest()
        {
//            Assert.AreEqual(0, settings.EmailCategories.Count);
            this.action.Perform();
//            Assert.AreNotEqual(0, settings.EmailCategories.Count);
        }
    }
}