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
using System.Net;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections;
using System.Windows.Forms;

namespace SuiteCRMClient
{
     
    public static class clsGlobals
    {        
        public static Uri SugarCRMURL { get; set; }
        public static HttpWebRequest CreateWebRequest(string url, int contentLength)
        {
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;

            request.Method = "POST";
            request.ContentLength = contentLength;
            request.ContentType = "application/x-www-form-urlencoded; charset=utf-8";
            return request;
        }

        public static string CreateFormattedPostRequest(string method, object parameters)
        {
            JsonSerializer serializer = new JsonSerializer();
            serializer.Converters.Add(new Newtonsoft.Json.Converters.JavaScriptDateTimeConverter());
            serializer.NullValueHandling = Newtonsoft.Json.NullValueHandling.Ignore;
            serializer.StringEscapeHandling = StringEscapeHandling.EscapeHtml;

            StringBuilder buffer = new StringBuilder();

            StringWriter swriter = new StringWriter(buffer);
            serializer.Serialize(swriter, parameters);
           
            string ret = "method=" + method;
            ret += "&input_type=JSON&response_type=JSON&rest_data=" + buffer.ToString();

            return ret;
        }

        public static T GetResponse<T>(string strMethod, object objInput, byte[] strFileContent = null)
        {
            try
            {
                string jsonData;
                jsonData = CreateFormattedPostRequest(strMethod, objInput);

                byte[] bytes = Encoding.UTF8.GetBytes(jsonData);

                HttpWebRequest request = CreateWebRequest(SugarCRMURL.AbsoluteUri + "service/v4_1/rest.php", bytes.Length);

                using (var requestStream = request.GetRequestStream())
                {
                    requestStream.Write(bytes, 0, bytes.Length);

                }
                using (var response = request.GetResponse() as HttpWebResponse)
                {
                    if (response.StatusCode != HttpStatusCode.OK)
                    {
                        clsSuiteCRMHelper.LoadLogFileLocation();
                        clsSuiteCRMHelper.AddLogLine("------------------" + System.DateTime.Now.ToString() + "-----------------");
                        clsSuiteCRMHelper.AddLogLine("GetResponse method Webserver Exception:");
                        clsSuiteCRMHelper.AddLogLine("Status Description:" + response.StatusDescription);
                        clsSuiteCRMHelper.AddLogLine("Status Code:" + response.StatusCode);
                        clsSuiteCRMHelper.AddLogLine("Method:" + response.Method);
                        clsSuiteCRMHelper.AddLogLine("Response URI:" + response.ResponseUri.ToString());
                        clsSuiteCRMHelper.AddLogLine("Inputs:");
                        clsSuiteCRMHelper.AddLogLine("Method:" + strMethod);
                        clsSuiteCRMHelper.AddLogLine("Data:" + objInput.ToString());
                        clsSuiteCRMHelper.AddLogLine("-------------------------------------------------------------------------");
                        clsSuiteCRMHelper.log.Close();
                        throw new Exception(response.StatusDescription);
                       
                    }
                    else
                    {
                        using (Stream input = response.GetResponseStream())
                        {
                            StreamReader reader = new StreamReader(input);
                            string buffer = reader.ReadToEnd();
                            var objReturn = JsonConvert.DeserializeObject<T>(buffer);
                            return objReturn;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                clsSuiteCRMHelper.LoadLogFileLocation();
                clsSuiteCRMHelper.AddLogLine("------------------" + System.DateTime.Now.ToString() + "-----------------");
                clsSuiteCRMHelper.AddLogLine("GetResponse method General Exception:");
                clsSuiteCRMHelper.AddLogLine("Message:" + ex.Message);
                clsSuiteCRMHelper.AddLogLine("Source:" + ex.Source);
                clsSuiteCRMHelper.AddLogLine("StackTrace:" + ex.StackTrace);
                clsSuiteCRMHelper.AddLogLine("Data:" + ex.Data.ToString());
                clsSuiteCRMHelper.AddLogLine("HResult:" + ex.HResult.ToString());
                clsSuiteCRMHelper.AddLogLine("Inputs:");
                clsSuiteCRMHelper.AddLogLine("Method:" + strMethod);
                clsSuiteCRMHelper.AddLogLine("Data:" + objInput.ToString());
                clsSuiteCRMHelper.AddLogLine("-------------------------------------------------------------------------");
                clsSuiteCRMHelper.log.Close();
                throw ex;
            }
        }
        
    }
}
