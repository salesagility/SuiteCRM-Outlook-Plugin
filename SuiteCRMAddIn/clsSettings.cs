/**
 * Outlook integration for SuiteCRM.
 * @package Outlook integration for SuiteCRM
 * @copyright SalesAgility Ltd http://www.salesagility.com
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU AFFERO GENERAL PUBLIC LICENSE as published by
 * the Free Software Foundation; either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU AFFERO GENERAL PUBLIC LICENSE
 * along with this program; if not, see http://www.gnu.org/licenses
 * or write to the Free Software Foundation,Inc., 51 Franklin Street,
 * Fifth Floor, Boston, MA 02110-1301  USA
 *
 * @author SalesAgility <info@salesagility.com>
 */
using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Specialized;
using System.Configuration;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

namespace SuiteCRMAddIn
{
    [CompilerGenerated, GeneratedCode("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "9.0.0.0")]
    public sealed class clsSettings : ApplicationSettingsBase
    {
        private static clsSettings defaultInstance = ((clsSettings)SettingsBase.Synchronized(new clsSettings()));

        public static clsSettings Default
        {
            get
            {
                return defaultInstance;
            }
        }

        [UserScopedSetting, DefaultSettingValue(""), DebuggerNonUserCode]
        public string username
        {
            get
            {
                return (string)this["username"];
            }
            set
            {
                this["username"] = value;
            }
        }
        [DefaultSettingValue(""), DebuggerNonUserCode, UserScopedSetting]
        public string password
        {
            get
            {
                return (string)this["password"];
            }
            set
            {
                this["password"] = value;
            }
        }
        [DefaultSettingValue(""), DebuggerNonUserCode, UserScopedSetting]
        public string host
        {
            get
            {
                return (string)this["host"];
            }
            set
            {
                this["host"] = value;
            }
        }
        [DefaultSettingValue("False"), UserScopedSetting, DebuggerNonUserCode]
        public bool ArchiveAttachmentsDefault
        {
            get
            {
                return (bool)this["ArchiveAttachmentsDefault"];
            }
            set
            {
                this["ArchiveAttachmentsDefault"] = value;
            }
        }
        [DebuggerNonUserCode, UserScopedSetting, DefaultSettingValue("False")]
        public bool AutomaticSearch
        {
            get
            {
                return (bool)this["AutomaticSearch"];
            }
            set
            {
                this["AutomaticSearch"] = value;
            }
        }
        [DefaultSettingValue("True"), DebuggerNonUserCode, UserScopedSetting]
        public bool ShowCustomModules
        {
            get
            {
                return (bool)this["ShowCustomModules"];
            }
            set
            {
                this["ShowCustomModules"] = value;
            }
        }
        [DefaultSettingValue("<?xml version=\"1.0\" encoding=\"utf-16\"?>\r\n<ArrayOfString xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">\r\n  <string>None|None</string>\r\n</ArrayOfString>"), UserScopedSetting, DebuggerNonUserCode]
        public StringCollection CustomModules
        {
            get
            {
                return (StringCollection)this["CustomModules"];
            }
            set
            {
                this["CustomModules"] = value;
            }
        }
        [DefaultSettingValue("1000"), DebuggerNonUserCode, UserScopedSetting]
        public int SyncMaxRecords
        {
            get
            {
                return (int)this["SyncMaxRecords"];
            }
            set
            {
                this["SyncMaxRecords"] = value;
            }
        }
        [DebuggerNonUserCode, UserScopedSetting, DefaultSettingValue("1,2,3")]
        public string SelectedSearchModules
        {
            get
            {
                return (string)this["SelectedSearchModules"];
            }
            set
            {
                this["SelectedSearchModules"] = value;
            }
        }
        [UserScopedSetting, DebuggerNonUserCode, DefaultSettingValue("False")]
        public bool ParticipateInCeip
        {
            get
            {
                return (bool)this["ParticipateInCeip"];
            }
            set
            {
                this["ParticipateInCeip"] = value;
            }
        }
        [DefaultSettingValue("True"), DebuggerNonUserCode, UserScopedSetting]
        public bool PopulateContextLookupList
        {
            get
            {
                return (bool)this["PopulateContextLookupList"];
            }
            set
            {
                this["PopulateContextLookupList"] = value;
            }
        }
        [DefaultSettingValue("True"), DebuggerNonUserCode, UserScopedSetting]
        public bool AutoArchive
        {
            get
            {
                return (bool)this["AutoArchive"];
            }
            set
            {
                this["AutoArchive"] = value;
            }
        }
        [DefaultSettingValue(""), DebuggerNonUserCode, UserScopedSetting]
        public System.Collections.Generic.List<string> AutoArchiveFolders
        {
            get
            {
                return (System.Collections.Generic.List<string>)this["AutoArchiveFolders"];               
            }
            set
            {
                this["AutoArchiveFolders"] = value;
            }
        }
        [UserScopedSetting, DefaultSettingValue(""), DebuggerNonUserCode]
        public string ExcludedEmails
        {
            get
            {
                return (string)this["ExcludedEmails"];
            }
            set
            {
                this["ExcludedEmails"] = value;
            }
        }
        [UserScopedSetting, DefaultSettingValue("True"), DebuggerNonUserCode]
        public bool IsFirstTime
        {
            get
            {
                return (bool)this["IsFirstTime"];
            }
            set
            {
                this["IsFirstTime"] = value;
            }
        }
      
        [DefaultSettingValue("False"), UserScopedSetting, DebuggerNonUserCode]
        public bool IsLDAPAuthentication
        {
            get
            {
                return (bool)this["IsLDAPAuthentication"];
            }
            set
            {
                this["IsLDAPAuthentication"] = value;
            }
        }
        [DefaultSettingValue(""), DebuggerNonUserCode, UserScopedSetting]
        public string LDAPKey
        {
            get
            {
                return (string)this["LDAPKey"];
            }
            set
            {
                this["LDAPKey"] = value;
            }
        }
        [DefaultSettingValue("True"), DebuggerNonUserCode, UserScopedSetting]
        public bool ShowConfirmationMessageArchive
        {
            get
            {
                return (bool)this["ShowConfirmationMessageArchive"];
            }
            set
            {
                this["ShowConfirmationMessageArchive"] = value;
            }
        }
        [DebuggerNonUserCode, UserScopedSetting]
        public StringCollection case_dropdown_priority
        {
            get
            {
                return (StringCollection)this["case_dropdown_priority"];
            }
            set
            {
                this["case_dropdown_priority"] = value;
            }
        }

        [UserScopedSetting, DebuggerNonUserCode]
        public StringCollection case_dropdown_status
        {
            get
            {
                return (StringCollection)this["case_dropdown_status"];
            }
            set
            {
                this["case_dropdown_status"] = value;
            }
        }

        public static Hashtable accountEntrys = new Hashtable();
        public static AutoCompleteStringCollection accounts = new AutoCompleteStringCollection();
    }
}
