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
namespace SuiteCRMAddIn.BusinessLogic.Tests
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass()]
    public class SyncDirectionTests
    {
        [TestMethod()]
        public void SyncDirectionToStringTest()
        {
            Assert.AreEqual("None", SyncDirection.ToString(SyncDirection.Direction.Neither));
            Assert.AreEqual("From CRM to Outlook", SyncDirection.ToString(SyncDirection.Direction.Export));
            Assert.AreEqual("From Outlook to CRM", SyncDirection.ToString(SyncDirection.Direction.Import));
            Assert.AreEqual("Both", SyncDirection.ToString(SyncDirection.Direction.BiDirectional));
        }

        [TestMethod()]
        public void SyncDirectionAllowOutboundTest()
        {
            Assert.IsTrue(SyncDirection.AllowOutbound(SyncDirection.Direction.BiDirectional), "Bidirectional includes both");
            Assert.IsTrue(SyncDirection.AllowOutbound(SyncDirection.Direction.Import), "Explicitly outbound");
            Assert.IsFalse(SyncDirection.AllowOutbound(SyncDirection.Direction.Export), "Explicitly not outbound");
            Assert.IsFalse(SyncDirection.AllowOutbound(SyncDirection.Direction.Neither), "Neither excludes both");
        }

        [TestMethod()]
        public void SyncDirectionAllowInboundTest()
        {
            Assert.IsTrue(SyncDirection.AllowInbound(SyncDirection.Direction.BiDirectional), "Bidirectional includes both");
            Assert.IsTrue(SyncDirection.AllowInbound(SyncDirection.Direction.Export), "Explicitly inbound");
            Assert.IsFalse(SyncDirection.AllowInbound(SyncDirection.Direction.Import), "Explicitly not inbound");
            Assert.IsFalse(SyncDirection.AllowInbound(SyncDirection.Direction.Neither), "Neither excludes both");
        }
    }
}