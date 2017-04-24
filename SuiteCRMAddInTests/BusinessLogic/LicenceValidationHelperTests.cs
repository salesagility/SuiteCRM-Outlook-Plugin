namespace SuiteCRMAddIn.BusinessLogic.Tests
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using SuiteCRMAddIn.BusinessLogic;
    using SuiteCRMAddIn.Tests;

    [TestClass()]
    public class LicenceValidationHelperTests : WithLoggerTests
    {
        [TestMethod()]
        public void ValidateTest()
        {
            Assert.IsTrue( new LicenceValidationHelper( this.Log,
                "b8794235718652747b82fd713deac078", "e10a9aff077e983deca51e5d3688636c").Validate(),
                "Key pair valid, should validate");
            Assert.IsFalse(new LicenceValidationHelper(this.Log,
                "b8794235718652747b82fd713deac078", "froboz").Validate(),
                "Customer licence key invalid, should not validate");
            Assert.IsFalse(new LicenceValidationHelper(this.Log,
                "froboz", "e10a9aff077e983deca51e5d3688636c").Validate(),
                "Public key invalid, should not validate");
        }
    }
}