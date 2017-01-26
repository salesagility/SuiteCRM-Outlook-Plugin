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
using System.Text;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using SuiteCRMClient.Logging;

namespace SuiteCRMClient
{
    public static class clsGlobals
    {
        private static readonly JsonSerializer Serializer;
        private static ILogger Log;

        static clsGlobals()
        {
            Serializer = new JsonSerializer();
            Serializer.Converters.Add(new Newtonsoft.Json.Converters.JavaScriptDateTimeConverter());
            Serializer.NullValueHandling = Newtonsoft.Json.NullValueHandling.Ignore;
        }

        public static void SetLog(ILogger log)
        {
            Log = log;
        }

        public static Uri SuiteCRMURL { get; set; }

        public static T GetCrmResponse<T>(string strMethod, object objInput, byte[] strFileContent = null, bool islog = false)
        {
            try
            {
                var request = CreateCrmRestRequest(strMethod, objInput);
                var buffer = GetResponseString(strMethod, objInput, islog, request);
                return JsonConvert.DeserializeObject<T>(buffer);
            }
            catch (Exception ex)
            {
                Log.Warn("Problem calling " + strMethod, ex);
                throw;
            }
        }

        private static HttpWebRequest CreateCrmRestRequest(string strMethod, object objInput)
        {
            var requestUrl = SuiteCRMURL.AbsoluteUri + "service/v4_1/rest.php";
            var restData = SerialiseJson(objInput);
            var jsonData = $"method={WebUtility.UrlEncode(strMethod)}&input_type=JSON&response_type=JSON&rest_data={WebUtility.UrlEncode(restData)}";

            var contentTypeAndEncoding = "application/x-www-form-urlencoded; charset=utf-8";
            var bytes = Encoding.UTF8.GetBytes(jsonData);
            return CreatePostRequest(requestUrl, bytes, contentTypeAndEncoding);
        }

        private static string SerialiseJson(object parameters)
        {
            var buffer = new StringBuilder();
            var swriter = new StringWriter(buffer);
            Serializer.Serialize(swriter, parameters);
            return buffer.ToString();
        }

        private static string GetResponseString(string strMethod, object objInput, bool islog, HttpWebRequest request)
        {
            using (var response = request.GetResponse() as HttpWebResponse)
            {
                if (response.StatusCode != HttpStatusCode.OK)
                {
                    LogFailedRequest(strMethod, objInput, response);
                    throw new Exception(response.StatusDescription);
                }

               return GetStringFromWebResponse(response, islog);
            }
        }

        private static string GetStringFromWebResponse(HttpWebResponse response, bool isLog)
        {
            using (var input = response.GetResponseStream())
            using(var reader = new StreamReader(input))
            {
                var result = reader.ReadToEnd();
                if (isLog)
                {
                    Log.Debug("Response : " + result);
                }
                return result;
            }
        }

        private static HttpWebRequest CreatePostRequest(string requestUrl, byte[] bytes, string contentTypeAndEncoding)
        {
            var request = WebRequest.Create(requestUrl) as HttpWebRequest;

            request.Method = "POST";
            request.ContentLength = bytes.Length;
            request.ContentType = contentTypeAndEncoding;

            using (var requestStream = request.GetRequestStream())
            {
                requestStream.Write(bytes, 0, bytes.Length);
            }
            return request;
        }

        private static void LogFailedRequest(string strMethod, object objInput, HttpWebResponse response)
        {
            Log.Warn(
                "GetResponse method Webserver Exception:" + "\n" +
                "Status Description:" + response.StatusDescription + "\n" +
                "Status Code:" + response.StatusCode + "\n"+
                "Method:" + response.Method + "\n"+
                "Response URI:" + response.ResponseUri.ToString() + "\n" +
                "Inputs:" + "\n"+
                "Method:" + strMethod + "\n"+
                "Data:" + objInput.ToString() + "\n" +
                "-------------------------------------------------------------------------" + "\n");
        }
    }
}
