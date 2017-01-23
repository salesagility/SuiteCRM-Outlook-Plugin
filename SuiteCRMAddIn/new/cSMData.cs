using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn
{
    [Serializable]
    public class cSMData
    {
        public string MEntryID { get; set; }
        public string SEntryID { get; set; }
        public string ModifiedDate { get; set; }
        public bool IsTouched { get; set; }
    }
        
    public static class SMDBHelper
    {
        public static clsSettings settings = Globals.ThisAddIn.settings;
        public static bool IsEntryIDExists(string sID)
        {
            settings.Reload();
            return settings.SMDB.Exists(a => a.MEntryID == sID);
        }

        public static void AddDBEntry(string sMID, string sSID, string sModifiedDate)
        {
            settings.SMDB.Add(new cSMData { MEntryID = sMID, SEntryID = sSID, ModifiedDate = sModifiedDate, IsTouched = true });
            settings.Save();
            settings.Reload();
        }

        public static bool IsUpdateRequiredOnUpdateEvent(string sID)
        {
            settings.Reload();
            return settings.SMDB.Exists(a => a.MEntryID == sID && a.ModifiedDate != "Fresh");
        }

        public static cSMData GetItem(string sSID)
        {
            settings.Reload();
            var oItem = settings.SMDB.Where(a => a.SEntryID == sSID).FirstOrDefault();
            return oItem;
        }

        public static object GetItemByID(string sID)
        {
            NameSpace oNS = Globals.ThisAddIn.Application.GetNamespace("mapi");
            return oNS.GetItemFromID(sID);
        }
    }
}
