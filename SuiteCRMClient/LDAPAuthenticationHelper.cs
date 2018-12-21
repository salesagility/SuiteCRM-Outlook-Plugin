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
    using RESTObjects;
    using System;
    using System.Globalization;
    using System.Security.Cryptography;
    using System.Text;

    /// <summary>
    /// Refactored out LDAP authentication to ease unit testing.
    /// </summary>
    public class LDAPAuthenticationHelper
    {
        private string username;
        private string password;
        private string key;

        /// <summary>
        /// The 'initialisation vector' - the encryption buffer gets initialised 
        /// to this before the password is written into it; not sure why.
        /// </summary>
        /// <remarks>
        /// Note that the initialisation vector is hardcoded as 'password' in the server
        /// side PHP code, so it probably isn't necessary to have it as a parameter here.
        /// </remarks>
        private string initialisationVector;

        private CrmRestServer server;

        private readonly string applicationName;

        /// <summary>
        /// Construct a new instance of LDAPAuthenticationHelper with these credentials.
        /// </summary>
        /// <param name="username">The username to identify as.</param>
        /// <param name="password">The pass to identify with.</param>
        /// <param name="key">The encryption key</param>
        /// <param name="iv">The initialization vector (in practice always "password").</param>
        /// <param name="service">The REST service exposed by the CRM instance.</param>
        public LDAPAuthenticationHelper( string username, string password, string key, string iv, string applicationName, CrmRestServer server)
        {
            this.username = username;
            this.password = password;
            this.key = key;
            this.initialisationVector = iv;
            this.applicationName = applicationName;
            this.server = server;
        }

        /// <summary>
        /// Authenticate me, via the CRM REST server, against LDAP with my credentials.
        /// </summary>
        /// <returns>The session id string on success, else a null or empty string.</returns>
        public string Authenticate()
        {
            string result = String.Empty;

            object loginData = new
            {
                user_auth = new
                {
                    user_name = username,
                    password = string.IsNullOrWhiteSpace(this.key) ?
                        this.password:
                        EncryptPassword(this.password,
                        ConstructEncryptionAlgorithm(this.key, this.initialisationVector)), 
                    encryption = string.IsNullOrWhiteSpace(this.key) ? 
                        "PLAIN" :
                        Boolean.TrueString,
                    application_name = this.applicationName
                }
            };

            return server.GetCrmResponse<RESTObjects.Login>("login", loginData).SessionID;
        }

        /// <summary>
        /// Construct an encryption algorithm using this key and this ?iv?.
        /// </summary>
        /// <param name="ldapKey">The LDAP key to use.</param>
        /// <param name="iv">The initialization vector to use (in practice, always "password").</param>
        /// <returns></returns>
        private SymmetricAlgorithm ConstructEncryptionAlgorithm(string ldapKey, string iv)
        {
            byte[] ldapKeyBuffer = new MD5CryptoServiceProvider().ComputeHash(Encoding.UTF8.GetBytes(ldapKey));
            StringBuilder ldapKeyBuilder = new StringBuilder();
            foreach (byte b in ldapKeyBuffer)
            {
                ldapKeyBuilder.Append(b.ToString("x2", CultureInfo.InvariantCulture));
            }
            TripleDES edes = new TripleDESCryptoServiceProvider
            {
                Mode = CipherMode.CBC,
                Key = Encoding.UTF8.GetBytes(ldapKeyBuilder.ToString(0, 0x18)),
                IV = Encoding.UTF8.GetBytes(iv),
                Padding = PaddingMode.Zeros
            };
            return edes;
        }

        /// <summary>
        /// Encrypt this pass with this algorithm.
        /// </summary>
        /// <param name="pass">The password to encrypt.</param>
        /// <param name="algie">The algorithm to encrypt it with.</param>
        /// <returns>The encrypted pass.</returns>
        private string EncryptPassword(string pass, SymmetricAlgorithm algie)
        {
            byte[] passwordBuffer = algie.CreateEncryptor().TransformFinalBlock(
                Encoding.UTF8.GetBytes(pass), 0,
                Encoding.UTF8.GetByteCount(pass));
            StringBuilder passwordBuilder = new StringBuilder();

            foreach (byte b in passwordBuffer)
            {
                passwordBuilder.Append(b.ToString("x2", CultureInfo.InvariantCulture));
            }

            return passwordBuilder.ToString();
        }
    }
}
