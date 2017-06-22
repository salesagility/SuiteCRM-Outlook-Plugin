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
    using System.Collections.Generic;
    using System.Collections;
    using SuiteCRMClient.Logging;

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
        /// The outlook item id of the item.
        /// </summary>
        public string OutlookId { get; set; }

        public object contactData;

        public UserSession SuiteCRMUserSession;

        public ArchiveableEmail(UserSession SuiteCRMUserSession, ILogger log)
        {
            this.log = log;
            this.SuiteCRMUserSession = SuiteCRMUserSession;
        }

        /// <summary>
        /// Get contact ids of all email addresses in my From, To, or CC fields which 
        /// are known to CRM and not included in these excluded addresses.
        /// </summary>
        /// <param name="contatenatedExcludedAddresses">A string containing zero or many email 
        /// addresses for which contact ids should not be returned.</param>
        /// <returns>A list of strings representing CRM contact ids.</returns>
        public List<string> GetValidContactIDs(string contatenatedExcludedAddresses = "")
        {
            return GetValidContactIds(ConstructAddressList(contatenatedExcludedAddresses.ToUpper()));
        }

        /// <summary>
        /// Get contact ids of all email addresses in my From, To, or CC fields which 
        /// are known to CRM and not included in these excluded addresses.
        /// </summary>
        /// <param name="excludedAddresses">email addresses for which contact ids should 
        /// not be returned.</param>
        /// <returns>A list of strings representing CRM contact ids.</returns>
        private List<string> GetValidContactIds(List<string> excludedAddresses)
        {
            RestAPIWrapper.EnsureLoggedIn(SuiteCRMUserSession);

            List<string> result = new List<string>();
            List<string> checkedAddresses = new List<string>();

            try
            {
                foreach (string address in ConstructAddressList($"{this.From};{this.To};{this.CC}"))
                {
                    if (!checkedAddresses.Contains(address) && !excludedAddresses.Contains(address.ToUpper()))
                    {
                        var contactReturn = SuiteCRMUserSession.RestServer.GetCrmResponse<RESTObjects.Contacts>("get_entry_list",
                            ConstructGetContactIdByAddressPacket(address));

                        if (contactReturn.entry_list != null && contactReturn.entry_list.Count > 0)
                        {
                            result.Add(contactReturn.entry_list[0].id);
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
            addresses.AddRange(contatenatedAddresses.Split(',', ';', '\n', ':', ' ', '\t'));
            return addresses;
        }

        private object ConstructGetContactIdByAddressPacket(string address)
        {
            return new
            {
                @session = SuiteCRMUserSession.id,
                @module_name = "Contacts",
                @query = GetContactIDQuery(address),
                @order_by = "",
                @offset = 0,
                @select_fields = new string[] { "id" },
                @max_results = 1
            };
        }

        private string GetContactIDQuery(string strEmail)
        {
            return "contacts.id in (SELECT eabr.bean_id FROM email_addr_bean_rel eabr JOIN email_addresses ea ON (ea.id = eabr.email_address_id) WHERE eabr.deleted=0 and ea.email_address = '" + strEmail + "')";
        }

        /// <summary>
        /// Save my email to CRM, if it relates to any valid contacts.
        /// </summary>
        /// <param name="excludedEmails">Emails of contacts with which it should not be related.</param>
        public ArchiveResult Save(string excludedEmails = "")
        {
            return Save(GetValidContactIDs(excludedEmails));
        }


        /// <summary>
        /// Save my email to CRM, and link it to these contact ids.
        /// </summary>
        /// <remarks>
        /// In the original code there were two entirely different ways of archiving emails; one did the trick of 
        /// trying first with the HTML body, and if that failed trying again with it empty. The other did not. 
        /// I have no idea whether there is a benefit of this two-attempt strategy.
        /// </remarks>
        /// <param name="crmContactIds">The contact ids to link with.</param>
        public ArchiveResult Save(List<string> crmContactIds)
        {
            ArchiveResult result;

            if (crmContactIds.Count > 0)
            {
                try
                {
                    result = TrySave(crmContactIds, this.HTMLBody, null);
                }
                catch (Exception firstFail)
                {
                    log.Warn($"ArchiveableEmail.Save: failed to save '{this.Subject}' with HTML body", firstFail);

                    try
                    {
                        result = TrySave(crmContactIds, string.Empty, new[] { firstFail });
                    }
                    catch (Exception secondFail)
                    {
                        log.Error($"ArchiveableEmail.Save: failed to save '{this.Subject}' at all", secondFail);
                        result = ArchiveResult.Failure(new[] { firstFail, secondFail });
                    }
                }
            }
            else
            {
                result = ArchiveResult.Failure(null);
            }
            
            return result;
        }


        /// <summary>
        /// Attempt to save me given these contact Ids and this HTML body, taking note of these previous failures.
        /// </summary>
        /// <param name="contactIds">CRM ids of contacts to which I should be related.</param>
        /// <param name="htmlBody">The HTML body with which I should be saved.</param>
        /// <param name="fails">Any previous failures in attempting to save me.</param>
        /// <returns>An archive result object describing the outcome of this attempt.</returns>
        private ArchiveResult TrySave(List<string> contactIds, string htmlBody, Exception[] fails)
        {
            var restServer = SuiteCRMUserSession.RestServer;
            var emailResult = restServer.GetCrmResponse<RESTObjects.SetEntryResult>("set_entry",
               ConstructPacket(htmlBody));
            ArchiveResult result = ArchiveResult.Success(emailResult.id, fails);

            SaveContacts(contactIds, emailResult);

            SaveAttachments(emailResult);

            return result;
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
                    BindAttachmentInCRM(emailResult.id,
                        TransmitAttachmentPacket(ConstructAttachmentPacket(attachment)).id);
                }
                catch (Exception any)
                {
                    log.Error($"Failed to bind attachment '{attachment.DisplayName}' to email '{emailResult.id}' in CRM", any);
                }
            }
        }

        /// <summary>
        /// Relate this email result (presumed to represent me) in CRM to these contact ids.
        /// </summary>
        /// <param name="crmContactIds">The contact ids which should be related to my email result.</param>
        /// <param name="emailResult">An email result (presumed to represent me).</param>
        private void SaveContacts(List<string> crmContactIds, RESTObjects.SetEntryResult emailResult)
        {
            var restServer = SuiteCRMUserSession.RestServer;

            foreach (string contactId in crmContactIds)
            {
                try
                {
                    restServer.GetCrmResponse<RESTObjects.eNewSetRelationshipListResult>("set_relationship",
                        ConstructContactRelationshipPacket(emailResult.id, contactId));
                }
                catch (Exception any)
                {
                    log.Error($"Failed to bind contact '{contactId}' to email '{emailResult.id}' in CRM", any);
                }
            }
        }

        /// <summary>
        /// Construct a packet representing my email.
        /// </summary>
        /// <returns>A packet which, when transmitted to CRM, will instantiate my email.</returns>
        private object ConstructPacket(string htmlBody)
        {
            List<RESTObjects.NameValue> emailData = new List<RESTObjects.NameValue>();
            emailData.Add(new RESTObjects.NameValue() { name = "from_addr", value = this.From });
            emailData.Add(new RESTObjects.NameValue() { name = "to_addrs", value = this.To.Replace("\n", "") });
            emailData.Add(new RESTObjects.NameValue() { name = "name", value = this.Subject });
            emailData.Add(new RESTObjects.NameValue() { name = "date_sent", value = this.Sent.ToString(EmailDateFormat) });
            emailData.Add(new RESTObjects.NameValue() { name = "description", value = this.Body });
            emailData.Add(new RESTObjects.NameValue() { name = "description_html", value = htmlBody });
            emailData.Add(new RESTObjects.NameValue() { name = "assigned_user_id", value = RestAPIWrapper.GetUserId() });
            emailData.Add(new RESTObjects.NameValue() { name = "status", value = "archived" });
            emailData.Add(new RESTObjects.NameValue() { name = "category_id", value = this.Category });

            object contactData = new
            {
                @session = SuiteCRMUserSession.id,
                @module_name = "Emails",
                @name_value_list = emailData
            };
            return contactData;
        }

        private void BindAttachmentInCRM(string emailId, string attachmentId)
        {
            //Relate the email and the attachment
            SuiteCRMUserSession.RestServer.GetCrmResponse<RESTObjects.eNewSetRelationshipListResult>("set_relationship",
                ConstructAttachmentRelationshipPacket(emailId, attachmentId));
        }

        /// <summary>
        /// Construct a packet representing the relationship between the email represented 
        /// by this email id and the attachment represented by this attachment id.
        /// </summary>
        /// <param name="emailId">The id of the email.</param>
        /// <param name="attachmentId">The id of the attachment.</param>
        /// <returns>A packet which, when transmitted to CRM, will instantiate this relationship.</returns>
        private object ConstructAttachmentRelationshipPacket(string emailId, string attachmentId)
        {
            return ConstructRelationshipPacket(emailId, attachmentId, "Emails", "notes");
        }

        /// <summary>
        /// Construct a packet representing the relationship between the email represented 
        /// by this email id and the contact represented by this contact id.
        /// </summary>
        /// <param name="emailId">The id of the email.</param>
        /// <param name="contactId">The id of the contact.</param>
        /// <returns>A packet which, when transmitted to CRM, will instantiate this relationship.</returns>
        private object ConstructContactRelationshipPacket(string emailId, string contactId)
        {
            return ConstructRelationshipPacket(contactId, emailId, "Contacts", "emails");
        }

        /// <summary>
        /// Construct a packet representing the relationship between the object represented 
        /// by this module id in the module with this module name and the object in the foreign
        /// module linked through this link field represented by this foreign id.
        /// </summary>
        /// <param name="moduleId">The id of the record in the named module.</param>
        /// <param name="foreignId">The id of the record in the foreign module.</param>
        /// <param name="moduleName">The name of the module in which the record is to be created.</param>
        /// <param name="linkField">The name of the link field in the named module which links to the foreign module.</param>
        /// <returns>A packet which, when transmitted to CRM, will instantiate this relationship.</returns>
        private object ConstructRelationshipPacket(string moduleId, string foreignId, string moduleName, string linkField)
        {
            return new
            {
                session = SuiteCRMUserSession.id,
                module_name = moduleName,
                module_id = moduleId,
                link_field_name = linkField,
                related_ids = new string[] { foreignId }
            };
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

            RESTObjects.NewNoteAttachment note = new RESTObjects.NewNoteAttachment();
            note.ID = res.id;
            note.FileName = attachment.DisplayName;
            note.FileCotent = attachment.FileContentInBase64String;

            object attachmentDataWebFormat = new
            {
                @session = SuiteCRMUserSession.id,
                @note = note
            };

            return attachmentDataWebFormat;
        }
    }
}
