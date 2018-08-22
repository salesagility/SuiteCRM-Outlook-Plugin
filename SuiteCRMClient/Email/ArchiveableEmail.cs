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
namespace SuiteCRMClient.Email
{
    using System;
    using System.Linq;
    using System.Collections.Generic;
    using SuiteCRMClient.Logging;
    using RESTObjects;

    /// <summary>
    /// A representation of an email which may be archived to CRM.
    /// </summary>
    public class ArchiveableEmail
    {
        /// <summary>
        /// Canonical format to use when saving date/times to CRM; essentially, ISO8601 without the 'T'.
        /// </summary>
        public const string EmailDateFormat = "yyyy-MM-dd HH:mm:ss";

        private readonly ILogger log;

        public string From { get; set; }
        public string To { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string HTMLBody { get; set; }
        public string CC { get; set; }

        public string Category { get; set; }

        /// <summary>
        /// Date/Time sent, if known, else now.
        /// </summary>
        public DateTime Sent { get; set; } = DateTime.UtcNow;

        public List<ArchiveableAttachment> Attachments { get; set; } = new List<ArchiveableAttachment>();
        public EmailArchiveReason Reason { get; set; }

        /// <summary>
        /// The client-side item id of the item.
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// The server-side item id.
        /// </summary>
        public string CrmEntryId { get; set; }

        public object contactData;

        public UserSession SuiteCRMUserSession;

        public ArchiveableEmail(UserSession SuiteCRMUserSession, ILogger log)
        {
            this.log = log;
            this.SuiteCRMUserSession = SuiteCRMUserSession;
        }

        /// <summary>
        /// Get related ids from the modules with these moduleKeys, of all email addresses 
        /// in my From, To, or CC fields which are known to CRM and not included in these 
        /// excluded addresses.
        /// </summary>
        /// <param name="moduleKeys">Keys (standardised names) of modules to search.</param>
        /// <param name="contatenatedExcludedAddresses">A string containing zero or many email 
        /// addresses for which contact ids should not be returned.</param>
        /// <returns>A list of pairs strings representing CRM module keys and contact ids.</returns>
        public IEnumerable<CrmEntity> GetRelatedIds(IEnumerable<string> moduleKeys, string contatenatedExcludedAddresses = "")
        {
            return GetRelatedIds(moduleKeys, ConstructAddressList(contatenatedExcludedAddresses.ToUpper()));
        }

        /// <summary>
        /// Get contact ids of all email addresses in my From, To, or CC fields which 
        /// are known to CRM and not included in these excluded addresses.
        /// </summary>
        /// <param name="moduleKeys">Keys (standardised names) of modules to search.</param>
        /// <param name="excludedAddresses">email addresses for which related ids should 
        /// not be returned.</param>
        /// <returns>A list of strings representing CRM ids.</returns>
        private IEnumerable<CrmEntity> GetRelatedIds(IEnumerable<string> moduleKeys, IEnumerable<string> excludedAddresses)
        {
            RestAPIWrapper.EnsureLoggedIn();

            List<CrmEntity> result = new List<CrmEntity>();
            List<string> checkedAddresses = new List<string>();

            try
            {
                foreach (string address in ConstructAddressList($"{this.From};{this.To};{this.CC}"))
                {
                    if (!checkedAddresses.Contains(address) && !excludedAddresses.Contains(address.ToUpper()))
                    {
                        RESTObjects.IdsOnly contacts = null;

                        foreach (string moduleKey in moduleKeys)
                        {
                            contacts = SuiteCRMUserSession.RestServer.GetCrmResponse<RESTObjects.IdsOnly>("get_entry_list",
                            ConstructGetContactIdByAddressPacket(address, moduleKey));


                            if (contacts.entry_list != null && contacts.entry_list.Count > 0)
                            {
                                result.AddRange(contacts.entry_list.Select(x => new CrmEntity(moduleKey, x.id)));
                            }
                        }
                    }
                    checkedAddresses.Add(address);
                }
            }
            catch (Exception ex)
            {
                log.Warn("GetValidContactIDs error", ex);
                throw;
            }

            return result;
        }

        /// <summary>
        /// From this concatenation of addresses, return a list of individual addresses.
        /// </summary>
        /// <remarks>
        /// There's no checking to establish that the addresses returned are even syntactically valid.
        /// </remarks>
        /// <param name="contatenatedAddresses">A string believed to contain 0 or many email 
        /// addresses separated by whitespace or punctuation.</param>
        /// <returns>A list of possible addresses extracted from the list.</returns>
        private static List<string> ConstructAddressList(string contatenatedAddresses)
        {
            List<string> addresses = new List<string>();
            addresses.AddRange(contatenatedAddresses
                .Split(',', ';', '\n', '\r', ':', ' ', '\t')
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim()));
            return addresses;
        }

        private object ConstructGetContactIdByAddressPacket(string address, string moduleKey)
        {
            string tableName = ModuleToTableResolver.GetTableName(moduleKey);
            return new
            {
                session = SuiteCRMUserSession.id,
                module_name = tableName,
                query = GetContactIDQuery(address, tableName),
                order_by = "",
                offset = 0,
                select_fields = new string[] { "id" },
                max_results = 1,
                deleted = false,
                favorites = false
            };
        }

        private string GetContactIDQuery(string address, string tableName)
        {
            return $"{tableName.ToLower()}.id in (SELECT eabr.bean_id FROM email_addr_bean_rel eabr JOIN email_addresses ea ON (ea.id = eabr.email_address_id) WHERE eabr.deleted=0 and ea.email_address = '{address}')";
        }

        /// <summary>
        /// Save my email to CRM, if it relates to any valid contacts.
        /// </summary>
        /// <param name="excludedEmails">Emails of contacts with which it should not be related.</param>
        /// <param name="moduleKeys">Keys (standardised names) of modules to search.</param>
        public ArchiveResult Save(IEnumerable<CrmEntity> relatedRecords, string excludedEmails = "")
        {
            IEnumerable<CrmEntity> withIds = relatedRecords.Where(x => !string.IsNullOrEmpty(x.EntityId));
            IEnumerable<CrmEntity> foundIds = GetRelatedIds(relatedRecords.Where(x => string.IsNullOrEmpty(x.EntityId)).Select(x => x.ModuleName), excludedEmails);
            return Save(withIds.Union(foundIds));
        }


        /// <summary>
        /// Save my email to CRM, and link it to these contact ids.
        /// </summary>
        /// <remarks>
        /// In the original code there were two entirely different ways of archiving emails; one did the trick of 
        /// trying first with the HTML body, and if that failed trying again with it empty. The other did not. 
        /// I have no idea whether there is a benefit of this two-attempt strategy.
        /// </remarks>
        /// <param name="relatedRecords">CRM module names/ids of records to which I should be related.</param>
        public ArchiveResult Save(IEnumerable<CrmEntity> relatedRecords)
        {
            ArchiveResult result;

            if (relatedRecords.Count() == 0)
            {
                result = ArchiveResult.Failure(
                    new[] { new Exception("Found no related entities in CRM to link with") });
            }
            else
            {
                try
                {
                    if (string.IsNullOrEmpty(this.CrmEntryId))
                    {
                        result = TrySave(relatedRecords, null);
                    }
                    else
                    {
                        result = TryUpdate(relatedRecords, null);
                    }
                }
                catch (Exception firstFail)
                {
                    log.Warn($"ArchiveableEmail.Save: failed to save '{this.Subject}' with HTML body", firstFail);

                    try
                    {
                        result = TrySave(relatedRecords, new[] { firstFail });
                    }
                    catch (Exception secondFail)
                    {
                        log.Error($"ArchiveableEmail.Save: failed to save '{this.Subject}' at all", secondFail);
                        result = ArchiveResult.Failure(new[] { firstFail, secondFail });
                    }
                }
            }
            
            return result;
        }


        /// <summary>
        /// Delete my existing record and create a new one.
        /// </summary>
        /// <param name="relatedRecords">CRM module names/ids of records to which I should be related.</param>
        /// <param name="fails">Any previous failures in attempting to save me.</param>
        /// <returns>An archive result object describing the outcome of this attempt.</returns>
        private ArchiveResult TryUpdate(IEnumerable<CrmEntity> relatedRecords, Exception[] fails)
        {
            ArchiveResult result;

            try
            {
                // delete
                NameValue[] deletePacket = new NameValue[2];
                deletePacket[0] = RestAPIWrapper.SetNameValuePair("id", this.CrmEntryId);
                deletePacket[1] = RestAPIWrapper.SetNameValuePair("deleted", "1");
                RestAPIWrapper.SetEntry(deletePacket, "Emails");
                // recreate
                result = this.TrySave(relatedRecords, fails);
            }
            catch (Exception any)
            {
                List<Exception> newFails = new List<Exception>();
                newFails.Add(any);
                if (fails != null && fails.Any())
                {
                    newFails.AddRange(fails);
                }
                result = ArchiveResult.Failure(newFails.ToArray());
            }

            return result;
        }

        /// <summary>
        /// Attempt to save me given these related records and this HTML body, taking note of these previous failures.
        /// </summary>
        /// <param name="relatedRecords">CRM module names/ids of records to which I should be related.</param>
        /// <param name="fails">Any previous failures in attempting to save me.</param>
        /// <returns>An archive result object describing the outcome of this attempt.</returns>
        private ArchiveResult TrySave(IEnumerable<CrmEntity> relatedRecords, Exception[] fails)
        {
            CrmRestServer restServer = SuiteCRMUserSession.RestServer;
            SetEntryResult emailResult = restServer.GetCrmResponse<RESTObjects.SetEntryResult>("set_entry",
               ConstructPacket(this.HTMLBody));
            ArchiveResult result = ArchiveResult.Success(emailResult.id, fails);

            if (result.IsSuccess)
            {
                LinkRelatedRecords(relatedRecords, emailResult);
                SaveAttachments(emailResult);
            }

            return result;
        }

        /// <summary>
        /// Relate this email result (presumed to represent me) in CRM to these related records.
        /// </summary>
        /// <param name="relatedRecords">The records which should be related to my email result.</param>
        /// <param name="emailResult">An email result (presumed to represent me).</param>
        private void LinkRelatedRecords(IEnumerable<CrmEntity> relatedRecords, RESTObjects.SetEntryResult emailResult)
        {
            var restServer = SuiteCRMUserSession.RestServer;

            foreach (CrmEntity record in relatedRecords)
            {
                try
                {
                    var success = RestAPIWrapper.TrySetRelationship(
                        new SetRelationshipParams
                        {
                            module2 = "emails",
                            module2_id = emailResult.id,
                            module1 = ModuleToTableResolver.GetTableName(record.ModuleName),
                            module1_id = record.EntityId,
                        }, Objective.Email);

                    if (success)
                    {
                        log.Debug($"Successfully bound {record.ModuleName} '{record.EntityId}' to email '{emailResult.id}' in CRM");
                    }
                    else
                    {
                        log.Warn($"Failed to bind {record.ModuleName} '{record.EntityId}' to email '{emailResult.id}' in CRM");
                    }
                }
                catch (Exception any)
                {
                    log.Error($"Failed to bind {record.ModuleName} '{record.EntityId}' to email '{emailResult.id}' in CRM", any);
                }
            }
        }

        /// <summary>
        /// Construct a packet representing my email.
        /// </summary>
        /// <returns>A packet which, when transmitted to CRM, will instantiate my email.</returns>
        private object ConstructPacket(string htmlBody)
        {
            EmailPacket emailData = new EmailPacket();

            emailData.MaybeAddField("from_addr_name", this.From);
            emailData.MaybeAddField("to_addrs_names", this.To, true);
            emailData.MaybeAddField("cc_addrs_names", this.CC, true);
            emailData.MaybeAddField("name", this.Subject);
            emailData.MaybeAddField("date_sent", this.Sent.ToString(EmailDateFormat));
            emailData.MaybeAddField("description", this.Body);
            emailData.MaybeAddField("description_html", htmlBody);
            emailData.MaybeAddField("assigned_user_id", RestAPIWrapper.GetUserId());
            emailData.MaybeAddField("category_id", this.Category);
            emailData.MaybeAddField("message_id", this.ClientId);

            return new
            {
                @session = SuiteCRMUserSession.id,
                @module_name = "Emails",
                @name_value_list = emailData
            };
        }


        /// <summary>
        /// Save my attachments to CRM, and relate them to this emailResult. 
        /// </summary>
        /// <param name="emailResult">A result object obtained by archiving me to CRM.</param>
        private void SaveAttachments(RESTObjects.SetEntryResult emailResult)
        {
            foreach (ArchiveableAttachment attachment in Attachments)
            {
                try
                {
                    BindAttachmentInCrm(emailResult.id,
                        TransmitAttachmentPacket(ConstructAttachmentPacket(attachment)).id);
                }
                catch (Exception any)
                {
                    log.Error($"Failed to bind attachment '{attachment.DisplayName}' to email '{emailResult.id}' in CRM", any);
                }
            }
        }


        /// <summary>
        /// Relate the email and the attachment
        /// </summary>
        /// <param name="emailId"></param>
        /// <param name="attachmentId"></param>
        /// <returns></returns>
        private bool BindAttachmentInCrm(string emailId, string attachmentId)
        {
            return RestAPIWrapper.TrySetRelationship(
                        new SetRelationshipParams
                        {
                            module2 = "emails",
                            module2_id = emailId,
                            module1 = "notes",
                            module1_id = attachmentId,
                        }, Objective.Email);
        }

 
        /// <summary>
        /// Transmit this attachment packet to CRM.
        /// </summary>
        /// <param name="attachmentPacket">The attachment packet to transmit</param>
        /// <returns>A result object indicating success or failure.</returns>
        private RESTObjects.SetEntryResult TransmitAttachmentPacket(object attachmentPacket)
        {
            return SuiteCRMUserSession.RestServer.GetCrmResponse<RESTObjects.SetEntryResult>("set_note_attachment", attachmentPacket);
        }

        /// <summary>
        /// Construct a packet representing this attachment.
        /// </summary>
        /// <remarks>
        /// Messy; confuses construction and transmission. Could do with refactoring.
        /// </remarks>
        /// <param name="attachment">The attachment to represent.</param>
        /// <returns>A packet which, when transmitted to CRM, will instantiate this attachment.</returns>
        private object ConstructAttachmentPacket(ArchiveableAttachment attachment)
        {
            List<RESTObjects.NameValue> initNoteData = new List<RESTObjects.NameValue>();
            initNoteData.Add(new RESTObjects.NameValue() { name = "name", value = attachment.DisplayName });

            object initNoteDataWebFormat = new
            {
                @session = SuiteCRMUserSession.id,
                @module_name = "Notes",
                @name_value_list = initNoteData
            };
            var res = SuiteCRMUserSession.RestServer.GetCrmResponse<RESTObjects.SetEntryResult>("set_entry", initNoteDataWebFormat);

            RESTObjects.NoteAttachment note = new RESTObjects.NoteAttachment();
            note.ID = res.id.ToString();
            note.FileName = attachment.DisplayName;
            note.FileContent = attachment.FileContentInBase64String;
            note.ParentType = "Emails";

            object attachmentDataWebFormat = new
            {
                @session = SuiteCRMUserSession.id,
                @note = note
            };

            return attachmentDataWebFormat;
        }

        private class EmailPacket : List<RESTObjects.NameValue>
        {
            public void MaybeAddField(string fieldName, string fieldValue, bool replaceCRs = false)
            {
                if (!string.IsNullOrWhiteSpace(fieldValue))
                {
                    this.Add(new RESTObjects.NameValue()
                    {
                        name = fieldName,
                        value = replaceCRs ? fieldValue.Replace("\n", "") : fieldValue
                    });
                }
            }

            public void MaybeAddField(string fieldName, string fieldValue)
            {
                this.MaybeAddField(fieldName, fieldValue.ToString(), false);
            }
        }
    }
}
