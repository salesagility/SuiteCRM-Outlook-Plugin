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
namespace SuiteCRMClient
{
    using System;
    using System.Text;
    using System.Security.Cryptography;
    using SuiteCRMClient.Logging;
    using Exceptions;

    public class UserSession
    {
        private readonly ILogger _log;

        public string SuiteCRMUsername { get; private set; }
        public string SuiteCRMPassword { get; private set; }
        public string LDAPKey { get; private set; }
        public string LDAPIV = "password";
        public bool AwaitingAuthentication { get; private set; } = true;

        private CrmRestServer restServer;

        /// <summary>
        /// The SuiteCRM session identifier.
        /// </summary>
        public string id { get; set; }

        public bool IsLoggedIn => !string.IsNullOrEmpty(id);

        public bool NotLoggedIn => !IsLoggedIn;

        public CrmRestServer RestServer
        {
            get { return this.restServer; }
        }

        public string ApplicationName { get; private set; }

        /// <summary>
        /// Construct a new instance of a UserSession. Note that all these parameters (except log) 
        /// come from the settings object, and it would be much simpler to just pass that in; 
        /// unfortunately, that's in the SuiteCRMAddIn assembly, and that is dependent on this, so
        /// can't be included. TODO: see if this could be refactored.
        /// </summary>
        /// <param name="URL">The URL of the rest handler to connect to.</param>
        /// <param name="Username">The username to authenticate as.</param>
        /// <param name="Password">The password to authenticate with.</param>
        /// <param name="ldapKey">The LDAP key to authenticate with.</param>
        /// <param name="log">The logger to log to.</param>
        /// <param name="timeout">The timeout for calls to the URL.</param>
        public UserSession(string URL, string Username, string Password, string ldapKey, string applicationName, ILogger log, int timeout)
        {
            _log = log;
            this.restServer = new CrmRestServer(log, timeout);

            if (URL != String.Empty)
            {
                this.restServer.SuiteCRMURL = new Uri(URL);
                this.SuiteCRMUsername = Username;
                this.SuiteCRMPassword = Password;
                this.LDAPKey = string.IsNullOrWhiteSpace(ldapKey) ? null : ldapKey;
                this.ApplicationName = applicationName;
            }
            id = String.Empty;
        }

        /// <summary>
        /// Logs in to the CRM server.
        /// </summary>
        /// <returns>if the server returned a 'polling_interval' value in the response packet, then that value, else null.</returns>
        public bool Login()
        {
            try
            {
                if (! String.IsNullOrWhiteSpace(LDAPKey))
                {
                    AuthenticateLDAP();
                }
                else
                {
                    AuthenticateCRM();
                }
            }
            catch (BadCredentialsException)
            {
                throw;
            }
            catch (Exception ex)
            {
                _log.Error("Login error", ex);
                id = String.Empty;
                SuiteCRMClient.RestAPIWrapper.SuiteCRMUserSession = null;
                throw;
            }
            finally
            {
                AwaitingAuthentication = false;
            }

            return this.IsLoggedIn;
        }

        /// <summary>
        /// Authenticate against CRM.
        /// </summary>
        /// <returns>A polling interval value, if returned by the host (currently it isn't)</returns>
        private void AuthenticateCRM()
        {
            AwaitingAuthentication = true;
            var username = SuiteCRMUsername != null ? SuiteCRMUsername : string.Empty;
            var password = this.SuiteCRMPassword != null ? this.SuiteCRMPassword : string.Empty;

            var hashedPass = global::SuiteCRMClient.UserSession.GetMD5Hash(password);
            RESTObjects.Login loginReturn;

            try
            {
                loginReturn = AuthenticateCRM(username, hashedPass);

                id = loginReturn.SessionID;
                SuiteCRMClient.RestAPIWrapper.SuiteCRMUserSession = this;
            }
            catch (BadCredentialsException)
            {
                throw;
            }
            catch (CrmServerErrorException e1)
            {
                _log.Error("Failed first login attempt", e1);
                try
                {
                    loginReturn = AuthenticateCRM(username, password);

                    id = loginReturn.SessionID;
                    SuiteCRMClient.RestAPIWrapper.SuiteCRMUserSession = this;

                }
                catch (BadCredentialsException)
                {
                    throw;
                }
                catch (CrmServerErrorException e2)
                {
                    _log.Error("Failed second login attempt", e2);
                    id = String.Empty;
                    SuiteCRMClient.RestAPIWrapper.SuiteCRMUserSession = null;
                    throw;
                }
            }
        }

        /// <summary>
        /// Authenticate against CRM, using this username and password.
        /// </summary>
        /// <param name="username">The username.</param>
        /// <param name="pass">The password.</param>
        /// <returns>An encapsulated session ID, essentially.</returns>
        private RESTObjects.Login AuthenticateCRM(string username, string pass)
        {
            object loginData = new
            {
                user_auth = new
                {
                    user_name = username,
                    password = pass,
                    application_name = this.ApplicationName
                }
            };
            return this.restServer.GetCrmResponse<RESTObjects.Login>("login", loginData);
        }

        /// <summary>
        /// Authenticate against LDAP using my configured credentials.
        /// </summary>
        private void AuthenticateLDAP()
        {
            try
            {
                this.AwaitingAuthentication = true;
                this.id = this.AuthenticateLDAP(this.SuiteCRMUsername, this.SuiteCRMPassword, this.LDAPKey, this.LDAPIV);
                if (String.IsNullOrWhiteSpace(this.id))
                {
                    this.id = String.Empty; // normalise away nulls
                    SuiteCRMClient.RestAPIWrapper.SuiteCRMUserSession = null;
                }
                else
                {
                    SuiteCRMClient.RestAPIWrapper.SuiteCRMUserSession = this;
                    this.AwaitingAuthentication = false;
                }
            }
            catch (Exception)
            {
                id = String.Empty;
                SuiteCRMClient.RestAPIWrapper.SuiteCRMUserSession = null;
                throw;
            }
        }

        /// <summary>
        /// Authenticate against the LDAP server ?implied by the SuiteCRM server? using 
        /// these credentials. Refactored out to assist unit testing.
        /// </summary>
        /// <param name="username">The username to authenticate.</param>
        /// <param name="password">The password which should be associated with this username.</param>
        /// <param name="key">The LDAP key entered by the user in the settings panel.</param>
        /// <param name="iv">?Purpose unknown, but value is always 'password'?</param>
        /// <returns></returns>
        public string AuthenticateLDAP(string username, string password, string key, string iv)
        {
            return new LDAPAuthenticationHelper(username, password, key, iv, ApplicationName,
                this.restServer).Authenticate();
        }

        /// <summary>
        /// The documentation at
        /// http://support.sugarcrm.com/Documentation/Sugar_Developer/Sugar_Developer_Guide_6.5/Application_Framework/Web_Services/Method_Calls/logout/
        /// says the response should be null, and it is. So that's fine. 
        /// </summary>
        public void LogOut()
        {
            try
            {
                if (! String.IsNullOrWhiteSpace( this.id))
                {
                    object logoutData = new
                    {
                        @session = this.id
                    };
                    var objRet = this.restServer.GetCrmResponse<object>("logout", logoutData);
                    _log.Debug("Logged out");
                }
            }
            catch (Exception ex)
            {
                _log.Error("Log out error", ex);
            }

            this.id = String.Empty;
            SuiteCRMClient.RestAPIWrapper.SuiteCRMUserSession = null;
        }

        public static string GetMD5Hash(string PlainText)
        {
            MD5 md = MD5.Create();
            byte[] bytes = Encoding.UTF8.GetBytes(PlainText);
            byte[] buffer2 = md.ComputeHash(bytes);
            StringBuilder builder = new StringBuilder(buffer2.Length);
            for (int i = 0; i < buffer2.Length; i++)
            {
                builder.Append(buffer2[i].ToString("X2"));
            }
            return builder.ToString();
        }

        /// <summary>
        /// AwaitingAuthentication was public write, which it really shouldn't 
        /// have been; I've fixed that, but it was set false in 
        /// <see cref="SuiteCRMAddIn.Dialogs.SettingsDialog#frmSettings_FormClosing"/>;
        /// I'm not certain why but it may have been important so I've
        /// (reluctantly) added this to enable it.
        /// </summary>
        public void ClearAwaitingAuthentication()
        {
            this.AwaitingAuthentication = false;
        }
    }
}
