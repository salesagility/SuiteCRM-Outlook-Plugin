/**
 * Outlook integration for SuiteCRM.
 * @package Outlook integration for SuiteCRM
 * @copyright SalesAgility Ltd http://www.salesagility.com
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU LESSER GENERAL PUBLIC LICENCE as published by
 * the Free Software Foundation; either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU LESSER GENERAL PUBLIC LICENCE
 * along with this program; if not, see http://www.gnu.org/licenses
 * or write to the Free Software Foundation,Inc., 51 Franklin Street,
 * Fifth Floor, Boston, MA 02110-1301  USA
 *
 * @author SalesAgility <info@salesagility.com>
 */
namespace SuiteCRMAddIn.BusinessLogic
{
    using System.Collections.Generic;
    using System.Linq;

    public class EmailAccountsArchiveSettings
    {
        public EmailAccountsArchiveSettings()
        {
            this.Clear();
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

        public void Load()
        {
            SelectedFolderEntryIds = new HashSet<string>(Properties.Settings.Default.AutoArchiveFolders ?? new List<string>());
            AccountsToArchiveInbound = new HashSet<string>(Properties.Settings.Default.AccountsToArchiveInbound ?? new List<string>());
            AccountsToArchiveOutbound = new HashSet<string>(Properties.Settings.Default.AccountsToArchiveOutbound ?? new List<string>());
        }

        public void Save()
        {
            Properties.Settings.Default.AutoArchiveFolders = SelectedFolderEntryIds.ToList();
            Properties.Settings.Default.AccountsToArchiveInbound = AccountsToArchiveInbound.ToList();
            Properties.Settings.Default.AccountsToArchiveOutbound = AccountsToArchiveOutbound.ToList();
        }

        public static EmailAccountsArchiveSettings Combine(IEnumerable<EmailAccountsArchiveSettings> accountSettings)
        {
            List<EmailAccountsArchiveSettings> settingsCopy = new List<EmailAccountsArchiveSettings>();
            settingsCopy.AddRange(accountSettings);

            var result = new EmailAccountsArchiveSettings();
            result.AccountsToArchiveInbound = new HashSet<string>(settingsCopy.SelectMany(s => s.AccountsToArchiveInbound));
            result.AccountsToArchiveOutbound = new HashSet<string>(settingsCopy.SelectMany(s => s.AccountsToArchiveOutbound));
            result.SelectedFolderEntryIds = new HashSet<string>(settingsCopy.SelectMany(s => s.SelectedFolderEntryIds));
            return result;
        }
    }
}
