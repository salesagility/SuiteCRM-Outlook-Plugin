namespace SuiteCRMClient.Tests
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using SuiteCRMAddInTests;
    using SuiteCRMAddInTests.Properties;
    using SuiteCRMClient;
    using System;

    [TestClass()]
    public class LDAPAuthenticationHelperTests : WithRestServiceTests
    {
        string validUser = Settings.Default.LDAPValidUser;
        string validPass = Settings.Default.LDAPValidPass;
        string validKey = Settings.Default.LDAPKey;

        [TestMethod()]
        public void LDAPAuthenticationHelperTest()
        {
            try
            {

                Assert.IsNotNull(new LDAPAuthenticationHelper(validUser, validPass, validKey, "password", "Unit Tests", server),
                    "Essentially all we need to test is that instantiation does not blow up");
            }
            catch (Exception)
            {
                Assert.Fail("Instantiation should not blow up");
            }
        }

        [TestMethod()]
        public void LDAPAuthenticationHelperAuthenticateTest()
        {
            Assert.IsFalse(
                String.IsNullOrWhiteSpace(
                    new LDAPAuthenticationHelper(validUser, validPass, validKey, "password", "Unit Tests", server).Authenticate()),
                "Good credentials, should validate.");
            Assert.IsTrue(
                String.IsNullOrWhiteSpace(
                    new LDAPAuthenticationHelper("invalid", validPass, validKey, "password", "Unit Tests", server).Authenticate()),
                "Bad username, should not validate");
            Assert.IsTrue(
                String.IsNullOrWhiteSpace(
                    new LDAPAuthenticationHelper(validUser, "invalid", validKey, "password", "Unit Tests", server).Authenticate()),
                "Bad password, should not validate");
            Assert.IsTrue(
                String.IsNullOrWhiteSpace(
                    new LDAPAuthenticationHelper(validUser, validPass, "invalid", "password", "Unit Tests", server).Authenticate()),
                    "Bad key, should not validate");
        }
    }
}