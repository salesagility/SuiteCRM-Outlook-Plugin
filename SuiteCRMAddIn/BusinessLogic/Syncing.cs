using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    using SuiteCRMClient.Logging;

    public class Syncing
    {
        private readonly SyncContext _context;

        public Syncing(SyncContext context)
        {
            _context = context;
        }

        protected SyncContext Context => _context;

        protected Outlook.Application Application => Context.Application;

        protected clsSettings settings => Context.settings;

        protected ILogger Log => Context.Log;

        public DateTime GetStartDate()
        {
            DateTime dtRet = DateTime.Now.AddDays(-5);
            return new DateTime(dtRet.Year, dtRet.Month, dtRet.Day, 0, 0, 0);
        }

        public string GetStartDateString()
        {
            return " AND [Start] >='" + GetStartDate().ToString("MM/dd/yyyy HH:mm") + "'";
        }
    }
}
