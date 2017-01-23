using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn
{
    public class cAppItem
    {        
        public string SEntryID { get; set; }
        public bool Touched { get; set; }
        public string SType { get; set; }
        public DateTime OModifiedDate { get; set; }
        public Outlook.AppointmentItem oItem { get; set; }
        public bool Delete { get; set; }
        private int _isupdate = 0;
        public int IsUpdate
        {
            get { return _isupdate; }
            set { _isupdate = value; }
        }
    }

    public class cTaskItem
    {
        public string SEntryID { get; set; }
        public bool Touched { get; set; }
        public DateTime OModifiedDate { get; set; }
        public Outlook.TaskItem oItem { get; set; }
        public bool Delete { get; set; }
        private int _isupdate = 0;
        public int IsUpdate
        {
            get { return _isupdate; }
            set { _isupdate = value; }
        }
    }
    public class cContactItem
    {
        public string SEntryID { get; set; }
        public bool Touched { get; set; }
        public DateTime OModifiedDate { get; set; }
        public Outlook.ContactItem oItem { get; set; }
        public bool Delete { get; set; }
        private int _isupdate = 0;
        public int IsUpdate {
            get { return _isupdate; }
            set { _isupdate = value; }
        }
    }
}
