namespace SuiteCRMAddIn.BusinessLogic.Tests
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Helpers;
    using SuiteCRMAddIn.Tests;
    using SuiteCRMAddInTests.Properties;

    [TestClass()]
    public class LicenceValidationHelperTests : WithLoggerTests
    {
        [TestMethod()]
        public void LicenceValidationHelperValidateTest()
        {
            Assert.IsTrue( new LicenceValidationHelper( this.Log,
                Settings.Default.LicencePublicKey, Settings.Default.LicenceCustomerKey).Validate(),
                "Key pair valid, should validate");
            Assert.IsFalse(new LicenceValidationHelper(this.Log,
                Settings.Default.LicencePublicKey, "invalid").Validate(),
                "Customer licence key invalid, should not validate");
            Assert.IsFalse(new LicenceValidationHelper(this.Log,
                "invalid", Settings.Default.LicenceCustomerKey).Validate(),
                "Public key invalid, should not validate");
        }
    }
}
