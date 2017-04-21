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
using System;
using System.Collections.Generic;
using System.Collections;
using SuiteCRMClient.Logging;

namespace SuiteCRMClient.Email
{
    /// <remarks>
    /// This class is truly horrid, and most of it is duplicated by better code in BusinessLogic/EmailArchiving.
    /// TODO: Refactor. See issue #125
    /// </remarks>
    public class clsEmailArchive
    {
        /// <summary>
        /// Canonical format to use when saving date/times to CRM; essentially, ISO8601 without the 'T'.
        /// </summary>
        public const string EmailDateFormat = "yyyy-MM-dd HH:mm:ss";

        private readonly ILogger _log;

        public string From { get; set; }
        public string To { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string HTMLBody { get; set; }
        public string CC { get; set; }

        /// <summary>
        /// Date/Time sent, if known, else now.
        /// </summary>
        public DateTime Sent { get; set; } = DateTime.UtcNow;

        public List<clsEmailAttachments> Attachments { get; set; } = new List<clsEmailAttachments>();
        public EmailArchiveType ArchiveType { get; set; }
        public object contactData;

        public UserSession SuiteCRMUserSession;

        public clsEmailArchive(UserSession SuiteCRMUserSession, ILogger log)
        {
            _log = log;
            this.SuiteCRMUserSession = SuiteCRMUserSession;
        }

        public clsEmailArchive()
        {
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
            clsSuiteCRMHelper.EnsureLoggedIn(SuiteCRMUserSession);

            List<string> result = new List<string>();
            List<string> checkedAddresses = new List<string>();

            try
            {
                foreach (string address in ConstructAddressList($"{From};{To};{CC}"))
                {
                    if (!checkedAddresses.Contains(address) && !excludedAddresses.Contains(address.ToUpper()))
                    {
                        var contactReturn = SuiteCRMUserSession.RestServer.GetCrmResponse<RESTObjects.eContacts>("get_entry_list",
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
                _log.Warn("GetValidContactIDs error", ex);
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

        /// <remarks>
        /// This is horrid. See See BusinessLogic.EmailArchiving.ConstructCrmItem; these need to be refactored together.
        /// TODO: Refactor. See issue #125
        /// </remarks>
        public void Save(string strExcludedEmails = "")
        {
            try
            {
                Save(GetValidContactIDs(strExcludedEmails));
            }
            catch (Exception ex)
            {
                _log.Error("clsEmailArchive.Save", ex);
                throw;
            }
        }

        /// <summary>
        /// Save my email to CRM, and link it to these contact ids.
        /// </summary>
        /// <param name="crmContactIds">The contact ids to link with.</param>
        public void Save(List<string> crmContactIds)
        {
            var restServer = SuiteCRMUserSession.RestServer;
            try
            {
                if (crmContactIds.Count > 0)
                {
                    var emailResult = restServer.GetCrmResponse<RESTObjects.eNewSetEntryResult>("set_entry",
                        ConstructEmailPacket());

                    foreach (string contactId in crmContactIds)
                    {
                        restServer.GetCrmResponse<RESTObjects.eNewSetRelationshipListResult>("set_relationship",
                            ConstructContactRelationshipPacket(emailResult.id, contactId));
                    }

                    foreach (clsEmailAttachments attachment in Attachments)
                    {
                        BindAttachmentInCRM(emailResult.id,
                            TransmitAttachmentPacket(ConstructAttachmentPacket(attachment)).id);
                    }
                }
            }
            catch (Exception ex)
            {
                _log.Error("clsEmailArchive.Save", ex);
                throw;
            }
        }

        /// <summary>
        /// Construct a packet representing my email.
        /// </summary>
        /// <returns>A packet which, when transmitted to CRM, will instantiate my email.</returns>
        private object ConstructEmailPacket()
        {
            List<RESTObjects.eNameValue> emailData = new List<RESTObjects.eNameValue>();
            emailData.Add(new RESTObjects.eNameValue() { name = "from_addr", value = From });
            emailData.Add(new RESTObjects.eNameValue() { name = "to_addrs", value = To.Replace("\n", "") });
            emailData.Add(new RESTObjects.eNameValue() { name = "name", value = Subject });
            emailData.Add(new RESTObjects.eNameValue() { name = "date_sent", value = Sent.ToString(EmailDateFormat) });
            emailData.Add(new RESTObjects.eNameValue() { name = "description", value = Body });
            emailData.Add(new RESTObjects.eNameValue() { name = "description_html", value = HTMLBody });
            emailData.Add(new RESTObjects.eNameValue() { name = "assigned_user_id", value = clsSuiteCRMHelper.GetUserId() });
            emailData.Add(new RESTObjects.eNameValue() { name = "status", value = "archived" });
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

        private RESTObjects.eNewSetEntryResult TransmitAttachmentPacket(object attachmentPacket)
        {
            return SuiteCRMUserSession.RestServer.GetCrmResponse<RESTObjects.eNewSetEntryResult>("set_note_attachment", attachmentPacket);
        }

        /// <summary>
        /// Construct a packet representing this attachment.
        /// </summary>
        /// <remarks>
        /// Messy; confuses construction and transmission. Could do with refactoring.
        /// </remarks>
        /// <param name="attachment">The attachment to represent.</param>
        /// <returns>A packet which, when transmitted to CRM, will instantiate this attachment.</returns>
        private object ConstructAttachmentPacket(clsEmailAttachments attachment)
        {
            List<RESTObjects.eNameValue> initNoteData = new List<RESTObjects.eNameValue>();
            initNoteData.Add(new RESTObjects.eNameValue() { name = "name", value = attachment.DisplayName });

            object initNoteDataWebFormat = new
            {
                @session = SuiteCRMUserSession.id,
                @module_name = "Notes",
                @name_value_list = initNoteData
            };
            var res = SuiteCRMUserSession.RestServer.GetCrmResponse<RESTObjects.eNewSetEntryResult>("set_entry", initNoteDataWebFormat);

            RESTObjects.eNewNoteAttachment note = new RESTObjects.eNewNoteAttachment();
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
