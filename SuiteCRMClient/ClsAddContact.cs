using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuiteCRMClient
{
    public class ClsAddContact
    {
        public string Salutation { get; set; }
        public string Firstname { get; set; }
        public string Lastname { get; set; }
        public string Title { get; set; }
        public string Department { get; set; }
        public string Description { get; set; }
        public string AccountName { get; set; }
        public string OfficePhone { get; set; }
        public string Mobile { get; set; }
        public string Fax { get; set; }

        public string  AssignedTo { get; set; }

        public string Primary_address_street{ get; set; }
        public string Primary_address_city { get; set; }
        public string Primary_address_state { get; set; }
        public string Primary_address_postalcode { get; set; }
        public string Primary_address_country { get; set; }

        public string AltPrimary_address_street { get; set; }
        public string AltPrimary_address_city { get; set; }
        public string AltPrimary_address_state { get; set; }
        public string AltPrimary_address_postalcode { get; set; }
        public string AltPrimary_address_country { get; set; }

        public string ReportTo { get; set; }
        public string LeadSource { get; set; }
        public string Campaign { get; set; }
        public bool SynctoOutlook { get; set; }
        public bool DoNotCall { get; set; }

        public bool OptedOut { get; set; }
        public bool Invaild { get; set; }
        public bool PrimaryEmail { get; set; }


        public string Email { get; set; }
    }
}
