using SuiteCRMClient.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SuiteCRMAddIn.BusinessLogic
{
    public class SyncContext
    {
        private readonly Outlook.Application _application;
        private readonly clsSettings _settings;
        private Outlook.OlItemType _currentFolderItemType;

        public SyncContext(Outlook.Application application, clsSettings settings)
        {
            _application = application;
            _settings = settings;
            _currentFolderItemType = Outlook.OlItemType.olMailItem;
        }

        public Outlook.Application Application => _application;

        public clsSettings settings => _settings;

        public ILogger Log => Globals.ThisAddIn.Log;

        public Outlook.OlItemType CurrentFolderItemType => _currentFolderItemType;

        public void SetCurrentFolder(Outlook.MAPIFolder folder)
        {
            _currentFolderItemType = folder.DefaultItemType;
        }
    }
}
