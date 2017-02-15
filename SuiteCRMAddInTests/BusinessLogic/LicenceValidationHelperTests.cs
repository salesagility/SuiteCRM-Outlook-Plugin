using Microsoft.VisualStudio.TestTools.UnitTesting;
using SuiteCRMAddIn.BusinessLogic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuiteCRMAddIn.BusinessLogic.Tests
{
    [TestClass()]
    public class LicenceValidationHelperTests
    {

        private IEnumerable<string> GetLogHeader()
        {
            yield return "Froboz";
        }

        [TestMethod()]
        public void ValidateTest()
        {
            Log4NetLogger logger = Log4NetLogger.FromFilePath("test", "C:\\temp\\suitecrmoutlook.log", () => GetLogHeader());
            Assert.IsTrue( new LicenceValidationHelper( logger,
                "b8794235718652747b82fd713deac078", "e10a9aff077e983deca51e5d3688636c").Validate(),
                "Key pair valid, should validate");
            Assert.IsFalse(new LicenceValidationHelper(logger,
                "b8794235718652747b82fd713deac078", "froboz").Validate(),
                "Customer licence key invalid, should not validate");
            Assert.IsFalse(new LicenceValidationHelper(logger,
                "froboz", "e10a9aff077e983deca51e5d3688636c").Validate(),
                "Public key invalid, should not validate");
        }
    }
}