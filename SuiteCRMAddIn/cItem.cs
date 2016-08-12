using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn
{
    public class cAppItem
    {        
        public string SEntryID { get; set; }
        public bool Touched { get; set; }
        public string SType { get; set; }
        public string OModifiedDate { get; set; }
        public Outlook.AppointmentItem oItem { get; set; }
        public bool Delete { get; set; }
        
    }

    public class cTaskItem
    {
        public string SEntryID { get; set; }
        public bool Touched { get; set; }
        public string OModifiedDate { get; set; }
        public Outlook.TaskItem oItem { get; set; }
        public bool Delete { get; set; }
    }
    public class cContactItem
    {
        public string SEntryID { get; set; }
        public bool Touched { get; set; }
        public string OModifiedDate { get; set; }
        public Outlook.ContactItem oItem { get; set; }
        public bool Delete { get; set; }

    }
}
