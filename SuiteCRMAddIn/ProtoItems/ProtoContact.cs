
namespace SuiteCRMAddIn.ProtoItems
{
    using System;
    using Extensions;
    using SuiteCRMClient;
    using SuiteCRMClient.RESTObjects;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using BusinessLogic;

    /// <summary>
    /// Broadly, a C# representation of a CRM contact.
    /// </summary>
    public class ProtoContact : ProtoItem<Outlook.ContactItem>
    {
        private readonly string Body;
        private readonly string BusinessAddressCity;
        private readonly string BusinessAddressCountry;
        private readonly string BusinessAddressPostalCode;
        private readonly string BusinessAddressState;
        private readonly string BusinessAddressStreet;
        private readonly string BusinessFaxNumber;
        private readonly string BusinessTelephoneNumber;
        private readonly string CompanyName;
        private readonly string Department;
        private readonly string Email1Address;
        private readonly string FirstName;
        private readonly string HomeTelephoneNumber;
        private readonly string JobTitle;
        private readonly string LastName;
        private readonly string MobileTelephoneNumber;
        private readonly string Title;
        private readonly bool isPublic;
        private readonly CrmId CrmEntryId;

        public override string Description
        {
            get
            {
                return $"{FirstName} {LastName} ({Email1Address})";
            }
        }

        public ProtoContact(Outlook.ContactItem olItem)
        {
            this.Body = olItem.Body;
            this.BusinessAddressCity = olItem.BusinessAddressCity;
            this.BusinessAddressCountry = olItem.BusinessAddressCountry;
            this.BusinessAddressPostalCode = olItem.BusinessAddressPostalCode;
            this.BusinessAddressState = olItem.BusinessAddressState;
            this.BusinessAddressStreet = olItem.BusinessAddressStreet;
            this.BusinessFaxNumber = olItem.BusinessFaxNumber;
            this.BusinessTelephoneNumber = olItem.BusinessTelephoneNumber;
            this.CompanyName = olItem.CompanyName;
            this.Department = olItem.Department;
            this.Email1Address = olItem.Email1Address;
            this.CrmEntryId = olItem.GetCrmId();
            this.FirstName = olItem.FirstName;
            this.HomeTelephoneNumber = olItem.HomeTelephoneNumber;
            this.JobTitle = olItem.JobTitle;
            this.LastName = olItem.LastName;
            this.MobileTelephoneNumber = olItem.MobileTelephoneNumber;
            this.Title = olItem.Title;
            this.isPublic = olItem.Sensitivity == Microsoft.Office.Interop.Outlook.OlSensitivity.olNormal;
        }

        public override NameValueCollection AsNameValues()
        {
            return new NameValueCollection
            {
                RestAPIWrapper.SetNameValuePair("email1", Email1Address),
                RestAPIWrapper.SetNameValuePair("title", JobTitle),
                RestAPIWrapper.SetNameValuePair("phone_work", BusinessTelephoneNumber),
                RestAPIWrapper.SetNameValuePair("phone_home", HomeTelephoneNumber),
                RestAPIWrapper.SetNameValuePair("phone_mobile", MobileTelephoneNumber),
                RestAPIWrapper.SetNameValuePair("phone_fax", BusinessFaxNumber),
                RestAPIWrapper.SetNameValuePair("department", Department),
                RestAPIWrapper.SetNameValuePair("primary_address_city", BusinessAddressCity),
                RestAPIWrapper.SetNameValuePair("primary_address_state", BusinessAddressState),
                RestAPIWrapper.SetNameValuePair("primary_address_postalcode", BusinessAddressPostalCode),
                RestAPIWrapper.SetNameValuePair("primary_address_country", BusinessAddressCountry),
                RestAPIWrapper.SetNameValuePair("primary_address_street", BusinessAddressStreet),
                RestAPIWrapper.SetNameValuePair("description", Body),
                RestAPIWrapper.SetNameValuePair("last_name", LastName),
                RestAPIWrapper.SetNameValuePair("first_name", FirstName),
                RestAPIWrapper.SetNameValuePair("account_name", CompanyName),
                RestAPIWrapper.SetNameValuePair("salutation", Title),
                CrmId.IsValid(CrmEntryId)
                    ? RestAPIWrapper.SetNameValuePair("id", CrmEntryId.ToString())
                    : RestAPIWrapper.SetNameValuePair("assigned_user_id", RestAPIWrapper.GetUserId()),
                RestAPIWrapper.SetNameValuePair("sync_contact", this.isPublic)
            };
        }
    }
}
