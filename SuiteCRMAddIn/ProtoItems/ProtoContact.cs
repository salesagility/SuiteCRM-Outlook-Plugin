
namespace SuiteCRMAddIn.ProtoItems
{ 
    using SuiteCRMClient;
    using SuiteCRMClient.RESTObjects;
    using System;
    using System.Collections.Generic;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Broadly, a C# representation of a CRM contact.
    /// </summary>
    public class ProtoContact : ProtoItem<Outlook.ContactItem>
    {
        private string Body;
        private string BusinessAddressCity;
        private string BusinessAddressCountry;
        private string BusinessAddressPostalCode;
        private string BusinessAddressState;
        private string BusinessAddressStreet;
        private string BusinessFaxNumber;
        private string BusinessTelephoneNumber;
        private string CompanyName;
        private string Department;
        private string Email1Address;
        private string FirstName;
        private string HomeTelephoneNumber;
        private string JobTitle;
        private string LastName;
        private string MobileTelephoneNumber;
        private string Title;

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
            this.FirstName = olItem.FirstName;
            this.HomeTelephoneNumber = olItem.HomeTelephoneNumber;
            this.JobTitle = olItem.JobTitle;
            this.LastName = olItem.LastName;
            this.MobileTelephoneNumber = olItem.MobileTelephoneNumber;
            this.Title = olItem.Title;
        }

        public override NameValueCollection AsNameValues(string entryId)
        {
            var data = new NameValueCollection();

            data.Add(RestAPIWrapper.SetNameValuePair("email1", Email1Address));
            data.Add(RestAPIWrapper.SetNameValuePair("title", JobTitle));
            data.Add(RestAPIWrapper.SetNameValuePair("phone_work", BusinessTelephoneNumber));
            data.Add(RestAPIWrapper.SetNameValuePair("phone_home", HomeTelephoneNumber));
            data.Add(RestAPIWrapper.SetNameValuePair("phone_mobile", MobileTelephoneNumber));
            data.Add(RestAPIWrapper.SetNameValuePair("phone_fax", BusinessFaxNumber));
            data.Add(RestAPIWrapper.SetNameValuePair("department", Department));
            data.Add(RestAPIWrapper.SetNameValuePair("primary_address_city", BusinessAddressCity));
            data.Add(RestAPIWrapper.SetNameValuePair("primary_address_state", BusinessAddressState));
            data.Add(RestAPIWrapper.SetNameValuePair("primary_address_postalcode", BusinessAddressPostalCode));
            data.Add(RestAPIWrapper.SetNameValuePair("primary_address_country", BusinessAddressCountry));
            data.Add(RestAPIWrapper.SetNameValuePair("primary_address_street", BusinessAddressStreet));
            data.Add(RestAPIWrapper.SetNameValuePair("description", Body));
            data.Add(RestAPIWrapper.SetNameValuePair("last_name", LastName));
            data.Add(RestAPIWrapper.SetNameValuePair("first_name", FirstName));
            data.Add(RestAPIWrapper.SetNameValuePair("account_name", CompanyName));
            data.Add(RestAPIWrapper.SetNameValuePair("salutation", Title));
            data.Add(string.IsNullOrEmpty(entryId) ?
                RestAPIWrapper.SetNameValuePair("assigned_user_id", RestAPIWrapper.GetUserId()) :
                RestAPIWrapper.SetNameValuePair("id", entryId));

            return data;
        }
    }
}
