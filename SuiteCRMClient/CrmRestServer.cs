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

using System.Diagnostics;

namespace SuiteCRMClient
{
    using Exceptions;
    using Newtonsoft.Json;
    using RESTObjects;
    using SuiteCRMClient.Logging;
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web;

    /// <summary>
    /// Low level communication with the REST server.
    /// </summary>
    /// <see cref="RestAPIWrapper"/> 
    public class CrmRestServer
    {
        private readonly JsonSerializer serialiser;
        private ILogger log;
        private int timeout = 0;

        /// <summary>
        /// It appears that CRM sends us back strings HTML escaped.
        /// </summary>
        private JsonSerializerSettings deserialiseSettings = new JsonSerializerSettings()
        {
            StringEscapeHandling = StringEscapeHandling.EscapeHtml
        };

        public CrmRestServer(ILogger log, int timeout)
        {
            this.log = log;
            this.timeout = timeout;
            serialiser = new JsonSerializer();
            serialiser.Converters.Add(new Newtonsoft.Json.Converters.JavaScriptDateTimeConverter());
            serialiser.NullValueHandling = Newtonsoft.Json.NullValueHandling.Ignore;
            serialiser.StringEscapeHandling = StringEscapeHandling.EscapeNonAscii;
        }

        public Uri SuiteCRMURL { get; set; }

        public string GetCrmStringResponse(string method, object input)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                HttpWebRequest request = CreateCrmRestRequest(method, input);
                string response = GetResponseString(request);
#if DEBUG
                LogRequest(request, method, input);
                LogResponse(response);
#endif
                CheckForCrmError(response, this.CreatePayload(method, input));

                if (response.StartsWith("\"") && response.EndsWith("\""))
                {
                    response = response.Substring(1, response.Length - 2);
                }

                return response;
            }
            catch (Exception ex)
            {
                log.Warn($"Tried calling '{method}' with parameter '{input}', timeout is {this.timeout}ms");
                log.Error($"Failed calling '{method}'", ex);
                throw;
            }
        }

        public T GetCrmResponse<T>(string method, object input)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                HttpWebRequest request = CreateCrmRestRequest(method, input);
                string jsonResponse = GetResponseString(request);
#if DEBUG
                LogRequest(request, method, input);
                LogResponse(jsonResponse);
#endif
                CheckForCrmError(jsonResponse, this.CreatePayload(method, input));

                return DeserializeJson<T>(jsonResponse);
            }
            catch (Exception ex)
            {
                log.Warn($"Tried calling '{method}' with parameter '{input}', timeout is {this.timeout}ms");
                log.Error($"Failed calling '{method}'", ex);
                throw;
            }
        }

        /// <summary>
        /// Send a get request with this path part to the server
        /// </summary>
        /// <param name="pathPart">The path part of the URL (may optionally include a query part)</param>
        /// <returns>True if the call resulted in a 200 status code.</returns>
        public bool SendGetRequest(string pathPart)
        {
            var requestUrl = SuiteCRMURL.AbsoluteUri + "/" + pathPart;
            HttpWebRequest request = WebRequest.Create(requestUrl) as HttpWebRequest;

            request.Method = "GET";
            request.Timeout = this.timeout;

            HttpWebResponse response = request.GetResponse() as HttpWebResponse;

            log.Info($"Sent `{requestUrl}`; received back status {response.StatusCode}");

            return (response.StatusCode == HttpStatusCode.OK);
        }

        /// <summary>
        /// Check whether this CRM response represents a CRM error, and if it does
        /// throw it as an exception.
        /// </summary>
        /// <param name="jsonResponse">A response from CRM.</param>
        /// <param name="payload">The payload of the request which gave rise to this response.</param>
        /// <exception cref="CrmServerErrorException">if the response was recognised as an error.</exception>
        private void CheckForCrmError(string jsonResponse, string payload)
        {
            ErrorValue error;
            try
            {
                error = DeserializeJson<ErrorValue>(jsonResponse);
            }
            catch (JsonSerializationException)
            {
                // it wasn't recognisable as an error. That's fine!
                error = new ErrorValue();
            }

            if (error != null && error.IsPopulated())
            {
                switch (Int32.Parse(error.number))
                {
                    case 10:
                    case 1008:
                    case 1009:
                        throw new BadCredentialsException(error);
                    default:
                        throw new CrmServerErrorException(error, HttpUtility.UrlDecode(payload));
                }
            }
        }

        private void LogResponse(string jsonResponse)
        {
            log.Debug($"Response from CRM: {jsonResponse}");
        }

        private void LogRequest(HttpWebRequest request, string method, object payload)
        {
            StringBuilder bob = new StringBuilder();
            bob.Append($"Request to CRM: \n\tURL: {request.RequestUri}\n\tMethod: {request.Method}\n");
            string content = CreatePayload(method, payload);
            bob.Append($"\tPayload: {content}\n");
            bob.Append($"\tDecoded: {HttpUtility.UrlDecode(content)}");
            log.Debug(bob.ToString());
        }

        private T DeserializeJson<T>(string responseJson)
        {
            try
            {
                return JsonConvert.DeserializeObject<T>(responseJson, deserialiseSettings);
            }
            catch (JsonReaderException parseError)
            {
                throw new Exception($"Failed to parse JSON ({parseError.Message}): {responseJson}");
            }
        }

        private HttpWebRequest CreateCrmRestRequest(string strMethod, object objInput)
        {
            try
            {
                var requestUrl = SuiteCRMURL.AbsoluteUri + "service/v4_1/rest.php";
                string jsonData = CreatePayload(strMethod, objInput);

                var contentTypeAndEncoding = "application/x-www-form-urlencoded; charset=utf-8";
                var bytes = Encoding.UTF8.GetBytes(jsonData);
#if DEBUG
                log.Debug($"CrmRestServer.CreateCrmRestRequest: data is {jsonData}");
                log.Debug($"CrmRestServer.CreateCrmRestRequest: bytes are {System.Text.Encoding.ASCII.GetString(bytes)}");
#endif
                return CreatePostRequest(requestUrl, bytes, contentTypeAndEncoding);
            }
            catch (Exception problem)
            {
                throw new Exception($"Could not construct '{strMethod}' request", problem);
            }
        }

        private string CreatePayload(string strMethod, object objInput)
        {
            var restData = SerialiseJson(objInput);
            var jsonData =
                $"method={WebUtility.UrlEncode(strMethod)}&input_type=JSON&response_type=JSON&rest_data={WebUtility.UrlEncode(restData)}";
            return jsonData;
        }

        private string SerialiseJson(object parameters)
        {
            var buffer = new StringBuilder();
            var swriter = new StringWriter(buffer);
            serialiser.Serialize(swriter, parameters);
            return buffer.ToString();
        }

        private string GetResponseString(HttpWebRequest request)
        {
            using (var response = request.GetResponse() as HttpWebResponse)
            {
                Debug.Assert(response != null, "response != null");
                if (response.StatusCode != HttpStatusCode.OK)
                {
                    throw new Exception($"{response.StatusCode} {response.StatusDescription} from {response.Method} {response.ResponseUri}");
                }

                return GetStringFromWebResponse(response);
            }
        }

        private string GetStringFromWebResponse(HttpWebResponse response)
        {
            using (var input = response.GetResponseStream())
            {
                Debug.Assert(input != null, "input != null");
                using (var reader = new StreamReader(input))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private HttpWebRequest CreatePostRequest(string requestUrl, byte[] bytes, string contentTypeAndEncoding)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            var request = WebRequest.Create(requestUrl) as HttpWebRequest;

            request.Method = "POST";
            request.ContentLength = bytes.Length;
            request.ContentType = contentTypeAndEncoding;
            request.Timeout = this.timeout;

            /* This block is really useful because it allows us to see exactly what gets sent over 
             * the wire, but it's also extremely dodgy because sensitive data will end up in the log.
             * It also puts a lot of clutter in the log! TODO: remove before stable release! */
#if DEBUG
            log.Debug(
                String.Format(
                    "CrmRestServer.CreatePostRequest:\n\tContent type: {0}\n\tPayload     {1}",
                    contentTypeAndEncoding,
                    System.Web.HttpUtility.UrlDecode(Encoding.ASCII.GetString(bytes).Trim())));
#endif

            using (var requestStream = request.GetRequestStream())
            {
                requestStream.Write(bytes, 0, bytes.Length);
            }
            return request;
        }
    }
}
