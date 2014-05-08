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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.Configuration;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using System.Windows.Forms;
using System.Collections;

namespace SuiteCRMClient
{

    public class clsUsersession
    {      
        public string SugarCRMUsername { get; set; }
        public string SugarCRMPassword { get; set; }
        public bool AwaitingAuthentication { get; set; }

        public string id { get; set; }
        
        public clsUsersession(string URL, string Username, string Password)
        {
            if (URL != "")
            {
                clsGlobals.SugarCRMURL = new Uri(URL);
                SugarCRMUsername = Username;
                SugarCRMPassword = Password;
            }
            id = "";            
        }

        public void Test()
        {
            object searchByModuleData = new
            {
                @session = id,
                @search_string = "SNT",
                @modules = new string[] { "Accounts", "Bugs", "Cases", "Contacts", "Leads", "Opportunities", "Project", "ProjectTask", "Quotes" },
                @max_results=1000,
                @assigned_user_id = "admin"              
            };
            var searchReturn = clsGlobals.GetResponse<RESTObjects.Login>("search_by_module", searchByModuleData);

        }

        public void Login()
        {
            try
            {
                AwaitingAuthentication = true;

                object loginData = new
                {
                    @user_auth = new
                    {
                        @user_name = SugarCRMUsername,
                        @password = GetMD5Hash(SugarCRMPassword, false)
                    }
                };
                var loginReturn = clsGlobals.GetResponse<RESTObjects.Login>("login", loginData);
                if (loginReturn.ErrorName != null)
                {
                    id = "";
                    throw new Exception(loginReturn.ErrorDescription);
                }
                else
                {
                    id = loginReturn.SessionID;
                    SuiteCRMClient.clsSuiteCRMHelper.SessionID = id;
                }

                AwaitingAuthentication = false;
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "clsUsersession.Login method General Exception:\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------\n";
                 clsSuiteCRMHelper.WriteLog(strLog);                
                throw ex;
            }
            
        }

        
        public void LogOut()
        {
            try
            {
                if (id != "")
                {
                    object logoutData = new
                    {
                        @session = id
                    };
                    var objRet = clsGlobals.GetResponse<object>("logout", logoutData);
                }
            }
            catch (Exception ex)
            {
                string strLog;
                strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                strLog += "clsUsersession.LogOut method General Exception:\n";
                strLog += "Message:" + ex.Message + "\n";
                strLog += "Source:" + ex.Source + "\n";
                strLog += "StackTrace:" + ex.StackTrace + "\n";
                strLog += "HResult:" + ex.HResult.ToString() + "\n";
                strLog += "-------------------------------------------------------------------------\n";                
                 clsSuiteCRMHelper.WriteLog(strLog);
                ex.Data.Clear();
            }
        }


        public static string GetMD5Hash(string value, bool upperCase)
        {
            // Instantiate new MD5 Service Provider to perform the hash
            System.Security.Cryptography.MD5CryptoServiceProvider md5ServiceProdivder = new System.Security.Cryptography.MD5CryptoServiceProvider();

            // Get a byte array representing the value to be hashed and hash it
            byte[] data = System.Text.Encoding.ASCII.GetBytes(value);
            data = md5ServiceProdivder.ComputeHash(data);

            // Get the hashed string value
            StringBuilder hashedValue = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
                hashedValue.Append(data[i].ToString("x2"));

            // Return the string in all caps if desired
            if (upperCase)
                return hashedValue.ToString().ToUpper();

            return hashedValue.ToString();
        }
    }
}
