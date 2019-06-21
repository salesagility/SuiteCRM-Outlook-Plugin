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
namespace SuiteCRMAddIn.Helpers
{
    using SuiteCRMClient;
    using SuiteCRMClient.Logging;
    using System;
    using System.Collections.Generic;
    using System.IO;
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

        public bool NeedNotRevlidate()
        {
            DateTime lastStart = Properties.Settings.Default.LVSLastStart;
            int startsRemaining = Properties.Settings.Default.LVSStartsRemaining;
            var daysSinceLastValidation = Math.Floor( DateTime.Now.Subtract(lastStart).TotalDays);

            bool result = // Properties.Settings.Default.LVSDisable ||
                (startsRemaining > 0 &&
                daysSinceLastValidation < Properties.Settings.Default.LVSPeriod);

            Properties.Settings.Default.LVSStartsRemaining--;
            Properties.Settings.Default.Save();

            return result;
        }

        /// <summary>
        /// Validate my key pair.
        /// </summary>
        /// <remarks>
        /// This program is open source. You are entitled to download the source and to make
        /// alterations. If you want to disable licence key checking, this method should just
        /// return true. However, we have put a great deal of work into writing this program
        /// for you, and want to continuing supporting it; so we'd appreciate it if you didn't.
        /// </remarks>
        /// <returns>true if validation succeeds, or if the validation server fails or times out; else false.</returns>
        public bool Validate()
        {
            /* Generally, assume that validation will fail. */
            bool result = NeedNotRevlidate();

            if (!result)
            {
                try
                {
                    try
                    {
                        IDictionary<string, string> parameters = new Dictionary<string, string>();
                        parameters["public_key"] = this.applicationKey;
                        parameters["key"] = this.licenceKey;

                        using (var response =
                            this.service.CreateGetRequest(parameters).GetResponse() as HttpWebResponse)
                        {
                            result = InterpretStatusCode(response);
                        }
                    }
                    catch (WebException badConnection)
                    {
                        logger.Error($"Failed to connect to licence server because {badConnection.Status}", badConnection);
                        switch (badConnection.Status)
                        {
                            case WebExceptionStatus.ProtocolError:
                                result = InterpretStatusCode((HttpWebResponse)badConnection.Response);
                                break;
                            case WebExceptionStatus.Timeout:
                                /* if the licence validation server fails to respond, treat that as OK */
                                result = true;
                                break;
                            case WebExceptionStatus.ConnectFailure:
                                /* if we can't connect, treat that as OK */
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
                finally
                {
                    if (result)
                    {
                        Properties.Settings.Default.LVSLastStart = DateTime.Now;
                        Properties.Settings.Default.LVSStartsRemaining = Properties.Settings.Default.LVSStarts;
                        Properties.Settings.Default.Save();
                    }
                }
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
                    /* the licence server doesn't actually report the encoding in a header, 
                     * but it seems to be ASCII. */
                    var encoding = Encoding.ASCII; 

                    using (var responseStream = response.GetResponseStream())
                    {
                        using (var reader = new StreamReader(responseStream, encoding))
                        {
                            logger.ShowAndAddEntry($"Licence server responded {reader.ReadToEnd()}", LogEntryType.Error);
                        }
                    }

                    result = false;
                    break;
                default:
                    logger.ShowAndAddEntry(
                        $"Licence server responded with an unexpected status code {response.StatusCode}", 
                        LogEntryType.Warning);
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
