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
using System.Collections;
using SuiteCRMClient.Logging;

namespace SuiteCRMClient.Email
{
    public class clsEmailArchive
    {
        private readonly ILogger _log;

        public string From { get; set; }
        public string To { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string HTMLBody { get; set; }
        public string CC { get; set; }
        public List<clsEmailAttachments> Attachments { get; set; }
        public EmailArchiveType ArchiveType { get; set; }
        public object contactData;

        public clsUsersession SuiteCRMUserSession;

        public clsEmailArchive(clsUsersession SuiteCRMUserSession, ILogger log)
        {
            _log = log;
            this.SuiteCRMUserSession = SuiteCRMUserSession;
            Attachments = new List<clsEmailAttachments>();
        }

        public clsEmailArchive()
        {
            Attachments = new List<clsEmailAttachments>();
        }
        private ArrayList GetValidContactIDs(string strExcludedEmails = "")
        {
            clsSuiteCRMHelper.EnsureLoggedIn(SuiteCRMUserSession);

            ArrayList arrRet = new ArrayList();
            ArrayList arrCheckedList = new ArrayList();
            string strEmails = "";
            strEmails = From + ";" + To;
            //if (ArchiveType == 1)
            //    strEmails = From + ";" + To;
            //else if (ArchiveType == 2)
            //    strEmails = From;
            //else if (ArchiveType == 3)
            //    strEmails = To;
            if (strEmails != "")
            {
                try
                {
                    foreach (string strEmail in strEmails.Split(';'))
                    {
                        if (arrCheckedList.Contains(strEmail))
                            continue;

                        // To check Excluded Emails
                        if (strExcludedEmails != "")
                        {
                            string strMails = strExcludedEmails;
                            string[] arrMails = strMails.Split(',', ';', '\n', ':', ' ', '\t');
                            foreach (string strSplitEmails in arrMails)
                            {
                                if (strEmail.Trim().ToUpper() == strSplitEmails.Trim().ToUpper())
                                {
                                    return new ArrayList();
                                }
                            }
                        }

                        contactData = new
                          {
                              @session = SuiteCRMUserSession.id,
                              @module_name = "Contacts",
                              @query = GetContactIDQuery(strEmail),
                              @order_by = "",
                              @offset = 0,
                              @select_fields = new string[] { "id" },
                              @max_results = 1
                          };
                        var contactReturn = clsGlobals.GetCrmResponse<RESTObjects.eContacts>("get_entry_list", contactData);

                        if (contactReturn.entry_list.Count > 0)
                            arrRet.Add(contactReturn.entry_list[0].id);
                        arrCheckedList.Add(strEmail);
                    }
                }
                catch (Exception ex)
                {
                    string strLog;
                    strLog = "------------------" + System.DateTime.Now.ToString() + "-----------------\n";
                    strLog += "clsEmailArchive.GetValidContactIDs method General Exception:\n";
                    strLog += "Message:" + ex.Message + "\n";
                    strLog += "Source:" + ex.Source + "\n";
                    strLog += "StackTrace:" + ex.StackTrace + "\n";
                    strLog += "HResult:" + ex.HResult.ToString() + "\n";
                    strLog += "Inputs:" + "\n";
                    strLog += "Data:" + contactData.ToString() + "\n";
                    strLog += "-------------------------------------------------------------------------\n";
                    _log.Warn(strLog);

                    throw ex;
                }
            }
            return arrRet;
        }

        private string GetContactIDQuery(string strEmail)
        {
            return "contacts.id in (SELECT eabr.bean_id FROM email_addr_bean_rel eabr JOIN email_addresses ea ON (ea.id = eabr.email_address_id) WHERE eabr.deleted=0 and ea.email_address = '" + strEmail + "')";
        }

        public void Save(string strExcludedEmails = "")
        {
            try
            {
                ArrayList arrCRMContacts = GetValidContactIDs(strExcludedEmails);

                if (arrCRMContacts.Count > 0)
                {
                    List<RESTObjects.eNameValue> emailData = new List<RESTObjects.eNameValue>();
                    emailData.Add(new RESTObjects.eNameValue() { name = "from_addr", value = From });
                    emailData.Add(new RESTObjects.eNameValue() { name = "to_addrs", value = To.Replace("\n", "") });
                    emailData.Add(new RESTObjects.eNameValue() { name = "name", value = Subject });
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
                    var emailResult = clsGlobals.GetCrmResponse<RESTObjects.eNewSetEntryResult>("set_entry", contactData);


                    foreach (string strContactID in arrCRMContacts)
                    {
                        object contacRelationshipData = new
                        {
                            @session = SuiteCRMUserSession.id,
                            @module_name = "Contacts",
                            @module_id = strContactID,
                            @link_field_name = "emails",
                            @related_ids = new string[] { emailResult.id }
                        };
                        var relResult = clsGlobals.GetCrmResponse<RESTObjects.eNewSetRelationshipListResult>("set_relationship", contacRelationshipData);

                    }

                    //Attachments
                    foreach (clsEmailAttachments objAttachment in Attachments)
                    {
                        //Initialize AddIn attachment
                        List<RESTObjects.eNameValue> initNoteData = new List<RESTObjects.eNameValue>();
                        initNoteData.Add(new RESTObjects.eNameValue() { name = "name", value = objAttachment.DisplayName });

                        object initNoteDataWebFormat = new
                        {
                            @session = SuiteCRMUserSession.id,
                            @module_name = "Notes",
                            @name_value_list = initNoteData
                        };
                        var res = clsGlobals.GetCrmResponse<RESTObjects.eNewSetEntryResult>("set_entry", initNoteDataWebFormat);

                        //upload the attachment  
                        RESTObjects.eNewNoteAttachment attachment = new RESTObjects.eNewNoteAttachment();
                        attachment.ID = res.id;
                        attachment.FileName = objAttachment.DisplayName;
                        attachment.FileCotent = objAttachment.FileContentInBase64String;

                        object attachmentDataWebFormat = new
                        {
                            @session = SuiteCRMUserSession.id,
                            @note = attachment
                        };

                        var attachmentResult = clsGlobals.GetCrmResponse<RESTObjects.eNewSetEntryResult>("set_note_attachment", attachmentDataWebFormat);

                        //Relate the email and the attachment
                        object contacRelationshipData = new
                        {
                            @session = SuiteCRMUserSession.id,
                            @module_name = "Emails",
                            @module_id = emailResult.id,
                            @link_field_name = "notes",
                            @related_ids = new string[] { attachmentResult.id }
                        };
                        var rel = clsGlobals.GetCrmResponse<RESTObjects.eNewSetRelationshipListResult>("set_relationship", contacRelationshipData);

                    }
                }
            }
            catch (Exception ex)
            {
                _log.Error("clsEmailArchive.Save", ex);
                throw;
            }
        }
    }
}
