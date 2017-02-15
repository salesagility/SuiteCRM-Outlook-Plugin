using Microsoft.VisualStudio.TestTools.UnitTesting;
using SuiteCRMClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuiteCRMClient.Tests
{
    [TestClass()]
    public class LDAPAuthenticationHelperTests
    {
        string validUser = "";
        string validPass = "";
        string validKey = "";

        [TestMethod()]
        public void LDAPAuthenticationHelperTest()
        {
            try
            {

                Assert.IsNotNull(new LDAPAuthenticationHelper(validUser, validPass, validKey, "password"),
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
                    new LDAPAuthenticationHelper(validUser, validPass, validKey, "password").Authenticate()),
                "Good credentials, should validate.");
            Assert.IsTrue(
                String.IsNullOrWhiteSpace(
                    new LDAPAuthenticationHelper("invalid", validPass, validKey, "password").Authenticate()),
                "Bad username, should not validate");
            Assert.IsTrue(
                String.IsNullOrWhiteSpace(
                    new LDAPAuthenticationHelper(validUser, "invalid", validKey, "password").Authenticate()),
                "Bad password, should not validate");
            Assert.IsTrue(
                String.IsNullOrWhiteSpace(
                    new LDAPAuthenticationHelper(validUser, validPass, "invalid", "password").Authenticate()),
                    "Bad key, should not validate");
        }
    }
}