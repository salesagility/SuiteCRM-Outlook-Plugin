/**
 * Outlook integration for SuiteCRM.
 * @package Outlook integration for SuiteCRM
 * @copyright SalesAgility Ltd http://www.salesagility.com
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU AFFERO GENERAL PUBLIC LICENSE as published by
 * the Free Software Foundation; either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU AFFERO GENERAL PUBLIC LICENSE
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
    using System.Globalization;
    using SuiteCRMClient.Logging;
    using SuiteCRMClient.RESTObjects;
    using System.Runtime.CompilerServices;

    public class UserSession
    {
        private readonly ILogger _log;

        public string SuiteCRMUsername { get; set; }
        public string SuiteCRMPassword { get; set; }
        public string LDAPKey { get; set; }
        public string LDAPIV = "password";
        public bool AwaitingAuthentication { get; set; }

        /// <summary>
        /// The SuiteCRM session identifier.
        /// </summary>
        public string id { get; set; }

        public bool IsLoggedIn => !string.IsNullOrEmpty(id);

        public bool NotLoggedIn => !IsLoggedIn;

        public UserSession(string URL, string Username, string Password, string strLDAPKey, ILogger log)
        {
            _log = log;
            if (URL != String.Empty)
            {
                CrmRestServer.SuiteCRMURL = new Uri(URL);
                SuiteCRMUsername = Username;
                SuiteCRMPassword = Password;
                LDAPKey = strLDAPKey;
            }
            id = String.Empty;
        }

        public void Login()
        {
            try
            {
                if (! String.IsNullOrWhiteSpace(LDAPKey))
                {
                    AuthenticateLDAP();
                }
                else
                {
                    AwaitingAuthentication = true;
                    object loginData = new
                    {
                        @user_auth = new
                        {
                            @user_name = SuiteCRMUsername,
                            @password = GetMD5Hash(SuiteCRMPassword)
                        }
                    };
                    var loginReturn = CrmRestServer.GetCrmResponse<RESTObjects.Login>("login", loginData);
                    if (loginReturn.ErrorName != null)
                    {
                        loginData = new
                        {
                            @user_auth = new
                            {
                                @user_name = SuiteCRMUsername,
                                @password = SuiteCRMPassword
                            }
                        };
                        loginReturn = CrmRestServer.GetCrmResponse<RESTObjects.Login>("login", loginData);
                        if (loginReturn.ErrorName != null)
                        {
                            id = String.Empty;
                            SuiteCRMClient.clsSuiteCRMHelper.SuiteCRMUserSession = null;
                            throw new Exception(loginReturn.ErrorDescription);
                        }
                        else
                        {
                            id = loginReturn.SessionID;
                            SuiteCRMClient.clsSuiteCRMHelper.SuiteCRMUserSession = this;
                        }
                    }
                    else
                    {
                        id = loginReturn.SessionID;
                        SuiteCRMClient.clsSuiteCRMHelper.SuiteCRMUserSession = this;
                    }
                    AwaitingAuthentication = false;
                }
            }
            catch (Exception ex)
            {
                _log.Error("Login error", ex);
                id = String.Empty;
                SuiteCRMClient.clsSuiteCRMHelper.SuiteCRMUserSession = null;
                throw;
            }

        }

        /// <summary>
        /// Authenticate against LDAP using my configured credentials.
        /// </summary>
        public void AuthenticateLDAP()
        {
            try
            {
                this.AwaitingAuthentication = true;
                this.id = this.AuthenticateLDAP(this.SuiteCRMUsername, this.SuiteCRMPassword, this.LDAPKey, this.LDAPIV);
                if (String.IsNullOrWhiteSpace(this.id))
                {
                    this.id = String.Empty; // normalise away nulls
                    SuiteCRMClient.clsSuiteCRMHelper.SuiteCRMUserSession = null;
                }
                else
                {
                    SuiteCRMClient.clsSuiteCRMHelper.SuiteCRMUserSession = this;
                    this.AwaitingAuthentication = false;
                }
            }
            catch (Exception)
            {
                id = String.Empty;
                SuiteCRMClient.clsSuiteCRMHelper.SuiteCRMUserSession = null;
                throw;
            }
        }

        /// <summary>
        /// Authenticate against the LDAP server ?implied by the SuiteCRM server? using 
        /// these credentials. Refactored out to assist unit testing.
        /// </summary>
        /// <param name="username">The username to authenticate.</param>
        /// <param name="password">The password which should be associated with this username.</param>
        /// <param name="key">?unknown?</param>
        /// <param name="iv">?Purpose unknown, but value is always 'password'?</param>
        /// <returns></returns>
        public string AuthenticateLDAP(string username, string password, string key, string iv)
        {
            return new LDAPAuthenticationHelper(username, password, key, iv).Authenticate();
        }

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
                    var objRet = CrmRestServer.GetCrmResponse<object>("logout", logoutData);
                }
            }
            catch (Exception ex)
            {
                _log.Error("Log out error", ex);
            }
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
    }
}
