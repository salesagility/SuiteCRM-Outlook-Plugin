namespace SuiteCRMClient.Tests
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using SuiteCRMAddIn;
    using SuiteCRMAddIn.Tests;
    using SuiteCRMAddInTests;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    [TestClass()]
    public class LDAPAuthenticationHelperTests : WithRestServiceTests
    {
        string validUser = "";
        string validPass = "";
        string validKey = "";

        [TestMethod()]
        public void LDAPAuthenticationHelperTest()
        {
            try
            {

                Assert.IsNotNull(new LDAPAuthenticationHelper(validUser, validPass, validKey, "password", service),
                    "Essentially all we need to test is that instantiation does not blow up");
            }
            catch (Exception)
            {
                Assert.Fail("Instantiation should not blow up");
            }
        }

        [TestMethod()]
        public void AuthenticateTest()
        {
            Assert.IsFalse(
                String.IsNullOrWhiteSpace(
                    new LDAPAuthenticationHelper(validUser, validPass, validKey, "password", service).Authenticate()),
                "Good credentials, should validate.");
            Assert.IsTrue(
                String.IsNullOrWhiteSpace(
                    new LDAPAuthenticationHelper("invalid", validPass, validKey, "password", service).Authenticate()),
                "Bad username, should not validate");
            Assert.IsTrue(
                String.IsNullOrWhiteSpace(
                    new LDAPAuthenticationHelper(validUser, "invalid", validKey, "password", service).Authenticate()),
                "Bad password, should not validate");
            Assert.IsTrue(
                String.IsNullOrWhiteSpace(
                    new LDAPAuthenticationHelper(validUser, validPass, "invalid", "password", service).Authenticate()),
                    "Bad key, should not validate");
        }
    }
}