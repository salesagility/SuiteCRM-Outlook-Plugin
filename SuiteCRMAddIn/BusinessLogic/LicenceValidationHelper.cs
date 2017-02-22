namespace SuiteCRMAddIn.BusinessLogic
{
    using Newtonsoft.Json;
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Text;

    /// <summary>
    /// A helper to do licence validation.  See documentation at https://store.suitecrm.com/selling/license-api.
    /// </summary>
    public class LicenceValidationHelper
    {
        /// <summary>
        /// The URL to which licence validation requests are made.
        /// </summary>
        private const String validationURL = "https://store.suitecrm.com/api/v1/key/validate";

        /// <summary>
        /// The logger to which I shall log.
        /// </summary>
        private ILogger logger;

        /// <summary>
        /// The 'public key', common to all instances of the add-in and shipped with the installer.
        /// </summary>
        private String applicationKey;

        /// <summary>
        /// The 'key'; the purchaser's licence key. Unique to this customer (but not necessarily to only this instance of the plugin).
        /// </summary>
        private String licenceKey;

        /// <summary>
        /// The service to which I despatch validation requests.
        /// </summary>
        private RestService service;

        /// <summary>
        /// Construct a new instance of LicenceValidationHelper with this application key and this licence key.
        /// </summary>
        /// <param name="logger">The logger to which this licence validation helper will log.</param>
        /// <param name="applicationKey">The 'public key', common to all instances of the add-in and shipped with the installer.</param>
        /// <param name="licenceKey">The 'key', unique to this customer and entered through the settings panel.</param>
        public LicenceValidationHelper(ILogger logger, String applicationKey, String licenceKey)
        {
            this.logger = logger;
            this.service = new RestService(validationURL, logger);
            this.applicationKey = applicationKey;
            this.licenceKey = licenceKey;
        }

        /// <summary>
        /// Validate my key pair.
        /// </summary>
        /// <returns>true if validation succeeds, or if the validation server fails or times out; else false.</returns>
        public bool Validate()
        {
            /* Generally, assume that validation will fail. */
            bool result = false;

            try
            {
                try
                {
                    IDictionary<string,string> parameters = new Dictionary<string, string>();
                    parameters["public_key"] = this.applicationKey;
                    parameters["key"] = this.licenceKey;

                    using (var response = this.service.CreateGetRequest(parameters).GetResponse() as HttpWebResponse)
                    {
                        result = InterpretStatusCode(response);
                    }
                }
                catch (WebException badConnection)
                {
                    switch (badConnection.Status)
                    {
                        case WebExceptionStatus.ProtocolError:
                            result = InterpretStatusCode((HttpWebResponse)badConnection.Response);
                            break;
                        case WebExceptionStatus.Timeout:
                            /* if the licence validation server fails to respond, treat that as OK */
                            result = true;
                            break;
                        case WebExceptionStatus.NameResolutionFailure:
                            /* if the licence validation server cannot be found, treat that as OK */
                            result = true;
                            break;
                        default:
                            throw;
                    }
                }
            }
            catch (Exception any)
            {
                this.logger.Error("LicenceValidationHelper.Validate ", any);
            }

            logger.Info(
                String.Format(
                    "LicenceValidationHelper.Validate: returning {0}", result));

            return result;
        }

        /// <summary>
        /// Interpret the status code returned by the licence validation API. 200 is good,
        /// 500 is bad but acceptable, 400 is not acceptable.
        /// </summary>
        /// <param name="response">A response assumed to be from the validation server.</param>
        /// <returns>True if validation accepted else false.</returns>
        private bool InterpretStatusCode(HttpWebResponse response)
        {
            bool result;

            switch (response.StatusCode)
            {
                case HttpStatusCode.OK:
                    /* if the licence validation server says OK, that's OK */
                    result = service.GetPayload<LicenceValidation>(response).validated;
                    break;
                case HttpStatusCode.InternalServerError:
                    /* if the licence validation server breaks, treat that as OK */
                    result = true;
                    break;
                case HttpStatusCode.BadRequest:
                    /* that's a conventionally signalled fail. */
                    result = false;
                    break;
                default:
                    logger.Warn(
                        String.Format(
                            "LicenceValidationHelper.InterpretStatusCode: Unexpected status code {0}", 
                            response.StatusCode));
                    result = false;
                    break;
            }

            logger.Info(
                String.Format(
                    "LicenceValidationHelper.InterpretStatusCode: status code {0}, returning {1}",
                    response.StatusCode, 
                    result));

            return result;
        }
    }
}
