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
    using Newtonsoft.Json;
    using SuiteCRMClient.Logging;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web;

    /// <summary>
    /// A generic handler for a REST service. Derived in large part from Andrew's 
    /// CrmRestServer, q.v., but that is static and I'm really unhappy about the 
    /// number of static classes in this design. In the long term I'd like to make 
    /// CrmRestServer a subclass of this class, but that is a major change I'm not
    /// yet ready to make.
    /// </summary>
    public class RestService
    {
        /// <summary>
        /// The base URL of the service I wrap. TODO: should perhaps be a URI object.
        /// </summary>
        private readonly string baseUrl;

        /// <summary>
        /// The logger through which I shall log.
        /// </summary>
        private readonly ILogger log;

        /// <summary>
        /// The JSON serialiser through which payloads are serailised/deserialised.
        /// </summary>
        private readonly JsonSerializer serializer;

        public RestService(string url, ILogger log)
        {
            this.serializer = new JsonSerializer();
            this.serializer.Converters.Add(new Newtonsoft.Json.Converters.JavaScriptDateTimeConverter());
            this.serializer.NullValueHandling = Newtonsoft.Json.NullValueHandling.Ignore;
            this.baseUrl = url;
            this.log = log;
        }

        public T GetResponse<T>(string strMethod, object objInput)
        {
            try
            {
                var request = CreateRestRequest(strMethod, objInput);
                var jsonResponse = GetResponseString(request);
                return DeserializeJson<T>(jsonResponse);
            }
            catch (Exception ex)
            {
                this.log.Warn($"Tried calling '{strMethod}' with parameter '{objInput}'");
                this.log.Error($"Failed calling '{strMethod}'", ex);
                throw;
            }
        }

        /// <summary>
        /// Get a response of type T from my base URL by appending these parameters.
        /// </summary>
        /// <typeparam name="T">The type of response I anticipate</typeparam>
        /// <param name="parameters">Name/value pairs to send</param>
        /// <returns>The object received.</returns>
        public T GetResponse<T>(IDictionary<String,String> parameters)
        {
            try
            {
                return DeserializeJson<T>(GetResponseString(CreateGetRequest(parameters)));
            }
            catch (Exception ex)
            {
                this.log.Error($"Failed calling '{this.baseUrl}'", ex);
                throw;
            }
        }

        /// <summary>
        /// Extract the payload from this web response.
        /// </summary>
        /// <typeparam name="T">The type of payload anticipated.</typeparam>
        /// <param name="response">The response believed to contain a payload of type T.</param>
        /// <returns>The payload extracted.</returns>
        public T GetPayload<T>(HttpWebResponse response)
        {
            return DeserializeJson<T>(GetStringFromWebResponse(response));
        }

        /// <summary>
        /// Create an HTTP 'GET' request by appending thise name/value pairs to my base URL.
        /// </summary>
        /// <param name="parameters">Name/value pairs to send.</param>
        /// <returns>A request object.</returns>
        public HttpWebRequest CreateGetRequest(IDictionary<string, string> parameters)
        {
            StringBuilder bob = new StringBuilder(this.baseUrl).Append("?");

            foreach (string key in parameters.Keys)
            {
                bob.Append(HttpUtility.UrlEncode(key))
                    .Append("=")
                    .Append(HttpUtility
                    .UrlEncode(parameters[key])).Append("&");
            }

            String requestUrl = bob.ToString();
            // trim off the terminal ampersand
            requestUrl = requestUrl.Substring(0, requestUrl.Length - 1);
            log.Debug($"RestService.CreateGetResuest: sending '{requestUrl}'");

            HttpWebRequest request = WebRequest.CreateHttp(requestUrl);
            return request;
        }

        private T DeserializeJson<T>(string responseJson)
        {
            try
            {
                return JsonConvert.DeserializeObject<T>(responseJson);
            }
            catch (JsonReaderException parseError)
            {
                throw new Exception($"Failed to parse JSON ({parseError.Message}): {responseJson}");
            }
        }

        /// <summary>
        /// Create a REST request encapsulating the object objInput. TODO: Note that 
        /// this is highly dependent on how SuiteCRM does things and is not yet generic.
        /// Further work needed.
        /// </summary>
        /// <param name="strMethod">Purpose unknown.</param>
        /// <param name="objInput">The object to send.</param>
        /// <returns></returns>
        private HttpWebRequest CreateRestRequest(string strMethod, object objInput)
        {
            try
            {
                var requestUrl = this.baseUrl;
                var restData = this.SerialiseJson(objInput);
                var jsonData =
                    $"method={WebUtility.UrlEncode(strMethod)}&input_type=JSON&response_type=JSON&rest_data={WebUtility.UrlEncode(restData)}";

                var contentTypeAndEncoding = "application/x-www-form-urlencoded; charset=utf-8";
                var bytes = Encoding.UTF8.GetBytes(jsonData);
                return this.CreatePostRequest(requestUrl, bytes, contentTypeAndEncoding);
            }
            catch (Exception problem)
            {
                throw new Exception($"Could not construct '{strMethod}' request", problem);
            }
        }

        /// <summary>
        /// Serailise this object as JSON.
        /// </summary>
        /// <param name="input">The object to serialise.</param>
        /// <returns>The serialisation.</returns>
        private string SerialiseJson(object input)
        {
            var buffer = new StringBuilder();
            this.serializer.Serialize(new StringWriter(buffer), input);
            return buffer.ToString();
        }

        /// <summary>
        /// Return the content of the response to this web request as a string.
        /// </summary>
        /// <param name="request">The request to send.</param>
        /// <returns>A string representation of the response content.</returns>
        /// <exception cref="Exception">if the response status was not 200 OK</exception>
        private string GetResponseString(HttpWebRequest request)
        {
            using (var response = request.GetResponse() as HttpWebResponse)
            {
                if (response.StatusCode != HttpStatusCode.OK)
                {
                    throw new Exception($"{response.StatusCode} {response.StatusDescription} from {response.Method} {response.ResponseUri}");
                }

                return GetStringFromWebResponse(response);
            }
        }
        /// <summary>
        /// Return the content of this web response as a string.
        /// </summary>
        /// <param name="response">The response.</param>
        /// <returns>The content.</returns>
        private string GetStringFromWebResponse(HttpWebResponse response)
        {
            using (var input = response.GetResponseStream())
            using (var reader = new StreamReader(input))
            {
                return reader.ReadToEnd();
            }
        }

        private HttpWebRequest CreatePostRequest(string requestUrl, byte[] bytes, string contentTypeAndEncoding)
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

    }
}