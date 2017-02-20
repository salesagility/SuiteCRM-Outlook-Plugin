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
    using Logging;
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
        private string iv;
        private RestService service;

        /// <summary>
        /// Construct a new instance of LDAPAuthenticationHelper with these credentials.
        /// </summary>
        /// <param name="username">The username to identify as.</param>
        /// <param name="password">The password to identify with.</param>
        /// <param name="key">The ?key?</param>
        /// <param name="iv">The ?ldapIV? (in practice always "password").</param>
        /// <param name="service">The REST service exposed by the CRM instance.</param>
        public LDAPAuthenticationHelper( string username, string password, string key, string iv, RestService service)
        {
            this.username = username;
            this.password = password;
            this.key = key;
            this.iv = iv;
            this.service = service;
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
                @user_auth = new
                {
                    @user_name = username,
                    @password = EncryptPassword(this.password, 
                    ConstructEncryptionAlgorithm(this.key, this.iv))
                }
            };

            return service.GetResponse<eSetEntryResult>("login", loginData).id;
        }

        /// <summary>
        /// Construct an encryption algorithm using this key and this ?iv?.
        /// </summary>
        /// <param name="key">The LDAP key to use.</param>
        /// <param name="iv">The ?iv? to use (in practice, always "password").</param>
        /// <returns></returns>
        private SymmetricAlgorithm ConstructEncryptionAlgorithm(string key, string iv)
        {
            byte[] ldapKeyBuffer = new MD5CryptoServiceProvider().ComputeHash(Encoding.UTF8.GetBytes(key));
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
        /// Encrypt this password with this algorithm.
        /// </summary>
        /// <param name="password">The password to encrypt.</param>
        /// <param name="algie">The algorithm to encrypt it with.</param>
        /// <returns>The encrypted password.</returns>
        private string EncryptPassword(string password, SymmetricAlgorithm algie)
        {
            byte[] passwordBuffer = algie.CreateEncryptor().TransformFinalBlock(
                Encoding.UTF8.GetBytes(password), 0,
                Encoding.UTF8.GetByteCount(password));
            StringBuilder passwordBuilder = new StringBuilder();

            foreach (byte b in passwordBuffer)
            {
                passwordBuilder.Append(b.ToString("x2", CultureInfo.InvariantCulture));
            }

            return passwordBuilder.ToString();
        }
    }
}
