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
namespace SuiteCRMAddIn.Tests
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using SuiteCRMAddInTests.Properties;

    /// <summary>
    /// Abstract superclass for a test class which needs a CRM connection.
    /// </summary>
    public abstract class AbstractWithCrmConnectionTest : WithLoggerTests
    {
        /// <summary>
        /// The session through which we authenticate to CRM.
        /// </summary>
        public UserSession SuiteCRMUserSession { get; private set; }

        /// <summary>
        /// Authenticate to CRM.
        /// </summary>
        /// <returns>true if authentication succeeded, else false.</returns>
        private bool Authenticate()
        {
            bool result = false;
            var settings = Settings.Default;

            try
            {
                if (settings.host != String.Empty)
                {
                    SuiteCRMUserSession =
                        new SuiteCRMClient.UserSession(
                            settings.host,
                            settings.username,
                            settings.password,
                            settings.LDAPKey,
                            ThisAddIn.AddInTitle,
                            Log,
                            settings.RestTimeout);
                    try
                    {
                        SuiteCRMUserSession.Login();

                        if (SuiteCRMUserSession.IsLoggedIn)
                        {
                            result = true;
                        }
                    }
                    catch (Exception any)
                    {
                        Log.Error("Failure while trying to authenticate to CRM", any);
                    }
                }
                else
                {
                    // We don't have a URL to connect to, dummy the connection.
                    SuiteCRMUserSession =
                        new SuiteCRMClient.UserSession(
                            String.Empty,
                            String.Empty,
                            String.Empty,
                            String.Empty,
                            ThisAddIn.AddInTitle,
                            Log,
                            settings.RestTimeout);
                }
            }
            catch (Exception ex)
            {
                Log.Error("ThisAddIn.Authenticate", ex);
            }

            return result;
        }

        /// <summary>
        /// Initialize() is called once during test execution before
        /// test methods in this test class are executed.
        /// </summary>
        [TestInitialize()]
        public virtual void Initialize()
        {
            Authenticate();
        }

        /// <summary>
        /// Cleanup() is called once during test execution after
        /// test methods in this class have executed unless
        /// this test class' Initialize() method throws an exception.
        /// </summary>
        [TestCleanup()]
        public virtual void Cleanup()
        {
            SuiteCRMUserSession.LogOut();
        }
    }
}