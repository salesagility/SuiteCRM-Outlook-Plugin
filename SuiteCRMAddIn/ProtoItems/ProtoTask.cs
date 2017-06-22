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

namespace SuiteCRMAddIn.ProtoItems
{
    using SuiteCRMClient;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Broadly, a C# representation of a CRM task.
    /// </summary>
    public class ProtoTask : ProtoItem<Outlook.TaskItem>
    {
        private Outlook.TaskItem oItem;
        private string body = String.Empty;
        private string dateStart = string.Empty, dateDue = string.Empty;
        private string description = String.Empty;

        private string priority;

        private string status;
        private string subject;

        public ProtoTask(Outlook.TaskItem oItem)
        {
            this.oItem = oItem;
            this.subject = this.oItem.Subject;

            if (oItem.Body != null)
            {
                body = oItem.Body.ToString();
                var times = this.ParseTimesFromTaskBody(body);
                if (times != null)
                {
                    DateTime utcStart = new DateTime();
                    DateTime utcDue = new DateTime();
                    utcStart = oItem.StartDate.ToUniversalTime();
                    if (oItem.DueDate != null)
                        utcDue = oItem.DueDate.ToUniversalTime();
                    utcDue = utcDue.Add(times[1]);

                    //check max date, date must has value !
                    if (utcStart.ToUniversalTime().Year < 4000)
                        dateStart = string.Format("{0:yyyy-MM-dd HH:mm:ss}", utcStart.ToUniversalTime());
                    if (utcDue.ToUniversalTime().Year < 4000)
                        dateDue = string.Format("{0:yyyy-MM-dd HH:mm:ss}", utcDue.ToUniversalTime());
                }
                else
                {
                    TakePeriodFromOutlookItem(oItem);
                }
            }
            else
            {
                TakePeriodFromOutlookItem(oItem);
            }

            if (!string.IsNullOrEmpty(body))
            {
                int lastIndex = body.LastIndexOf("#<");
                if (lastIndex >= 0)
                    description = body.Remove(lastIndex);
                else
                {
                    description = body;
                }
            }

            switch (oItem.Status)
            {
                case Outlook.OlTaskStatus.olTaskNotStarted:
                    status = "Not Started";
                    break;
                case Outlook.OlTaskStatus.olTaskInProgress:
                    status = "In Progress";
                    break;
                case Outlook.OlTaskStatus.olTaskComplete:
                    status = "Completed";
                    break;
                case Outlook.OlTaskStatus.olTaskDeferred:
                    status = "Deferred";
                    break;
                default:
                    status = string.Empty;
                    break;
            }

            switch (oItem.Importance)
            {
                case Outlook.OlImportance.olImportanceLow:
                    priority = "Low";
                    break;

                case Outlook.OlImportance.olImportanceNormal:
                    priority = "Medium";
                    break;

                case Outlook.OlImportance.olImportanceHigh:
                    priority = "High";
                    break;
                default:
                    priority = string.Empty;
                    break;
            }

        }

        private void TakePeriodFromOutlookItem(Outlook.TaskItem oItem)
        {
            dateStart = oItem.StartDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
            dateDue = oItem.DueDate.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss");
        }

        /// <summary>
        /// True if other is also a ProtoTask I have identically the same content as other.
        /// </summary>
        /// <param name="other">Another object, which may be a prototask.</param>
        /// <returns>True if other is also a ProtoTask I have identically the same content as other.</returns>
        public override bool Equals(object other)
        {
            bool result = false;
            if (other is ProtoTask)
            {
                Dictionary<string, object> myContents = this.AsNameValues(string.Empty).AsDictionary();
                Dictionary<string, object> theirContents = ((ProtoTask)other).AsNameValues(string.Empty).AsDictionary();

                result = myContents.Keys.Count == theirContents.Keys.Count;
                foreach (string key in myContents.Keys)
                {
                    result &= myContents[key].Equals(theirContents[key]);
                }
            }

            return result;
        }

        /// <summary>
        /// I'm very like a dictionary constructed from my names/values, but not quite.
        /// </summary>
        /// <returns>A hash code </returns>
        public override int GetHashCode()
        {
            return this.AsNameValues(string.Empty).AsDictionary().GetHashCode() + 1;
        }

        /// <summary>
        /// Construct a name value list (to be serialised as JSON) representing this task.
        /// </summary>
        /// <param name="entryId">The presumed id of this task in CRM, if known.</param>
        /// <returns>a name value list representing this task</returns>
        public override NameValueCollection AsNameValues(string entryId)
        {
            var data = new NameValueCollection();
            data.Add(RestAPIWrapper.SetNameValuePair("name", this.subject));
            data.Add(RestAPIWrapper.SetNameValuePair("description", this.description));
            data.Add(RestAPIWrapper.SetNameValuePair("status", this.status));
            data.Add(RestAPIWrapper.SetNameValuePair("date_due", this.dateDue));
            data.Add(RestAPIWrapper.SetNameValuePair("date_start", this.dateStart));
            data.Add(RestAPIWrapper.SetNameValuePair("priority", this.priority));

            data.Add(String.IsNullOrEmpty(entryId) ?
                RestAPIWrapper.SetNameValuePair("assigned_user_id", RestAPIWrapper.GetUserId()) :
                RestAPIWrapper.SetNameValuePair("id", entryId));
            return data;
        }

        private TimeSpan[] ParseTimesFromTaskBody(string body)
        {
            try
            {
                if (string.IsNullOrEmpty(body))
                    return null;
                TimeSpan[] result = new TimeSpan[2];
                List<int> hhmm = new List<int>(4);

                string times = body.Substring(body.LastIndexOf("#<")).Substring(2);
                char[] sep = { '<', '#', ':' };
                int parsed = 0;
                foreach (var fragment in times.Split(sep))
                {
                    int.TryParse(fragment, out parsed);
                    hhmm.Add(parsed);
                    parsed = 0;
                }

                TimeSpan start_time = TimeSpan.FromHours(hhmm[0]).Add(TimeSpan.FromMinutes(hhmm[1]));
                TimeSpan due_time = TimeSpan.FromHours(hhmm[2]).Add(TimeSpan.FromMinutes(hhmm[3]));
                result[0] = start_time;
                result[1] = due_time;
                return result;
            }
            catch
            {
                // Log.Warn("Body doesn't have time string");
                return null;
            }
        }
    }
}
