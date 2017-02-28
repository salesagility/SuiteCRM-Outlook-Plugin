using System.Collections.Generic;
using System.Linq;

namespace SuiteCRMAddIn.BusinessLogic
{
    public class EmailAccountsArchiveSettings
    {
        public EmailAccountsArchiveSettings()
        {
        }

        public HashSet<string> SelectedFolderEntryIds { get; private set; }
        public HashSet<string> AccountsToArchiveInbound { get; private set; }
        public HashSet<string> AccountsToArchiveOutbound { get; private set; }

        public bool HasAny
            => SelectedFolderEntryIds.Count + AccountsToArchiveOutbound.Count + AccountsToArchiveInbound.Count != 0;

        public EmailAccountsArchiveSettings Clear()
        {
            SelectedFolderEntryIds = new HashSet<string>();
            AccountsToArchiveInbound = new HashSet<string>();
            AccountsToArchiveOutbound = new HashSet<string>();
            return this;
        }

        public void Load(clsSettings settings)
        {
            SelectedFolderEntryIds = new HashSet<string>(settings.AutoArchiveFolders);
            AccountsToArchiveInbound = new HashSet<string>(settings.AccountsToArchiveInbound);
            AccountsToArchiveOutbound = new HashSet<string>(settings.AccountsToArchiveOutbound);
        }

        public void Save(clsSettings settings)
        {
            settings.AutoArchiveFolders = SelectedFolderEntryIds.ToList();
            settings.AccountsToArchiveInbound = AccountsToArchiveInbound.ToList();
            settings.AccountsToArchiveOutbound = AccountsToArchiveOutbound.ToList();
        }

        public static EmailAccountsArchiveSettings Combine(IEnumerable<EmailAccountsArchiveSettings> accountSettings)
        {
            var result = new EmailAccountsArchiveSettings();
            result.AccountsToArchiveInbound = new HashSet<string>(accountSettings.SelectMany(s => s.AccountsToArchiveInbound));
            result.AccountsToArchiveOutbound = new HashSet<string>(accountSettings.SelectMany(s => s.AccountsToArchiveOutbound));
            result.SelectedFolderEntryIds = new HashSet<string>(accountSettings.SelectMany(s => s.SelectedFolderEntryIds));
            return result;
        }
    }
}
