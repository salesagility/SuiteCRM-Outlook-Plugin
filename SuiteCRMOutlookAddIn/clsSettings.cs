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

namespace SuiteCRMOutlookAddIn
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
        public bool archive_attachments_default
        {
            get
            {
                return (bool)this["archive_attachments_default"];
            }
            set
            {
                this["archive_attachments_default"] = value;
            }
        }
        [DebuggerNonUserCode, UserScopedSetting, DefaultSettingValue("False")]
        public bool automatic_search
        {
            get
            {
                return (bool)this["automatic_search"];
            }
            set
            {
                this["automatic_search"] = value;
            }
        }
        [DefaultSettingValue("True"), DebuggerNonUserCode, UserScopedSetting]
        public bool show_custom_modules
        {
            get
            {
                return (bool)this["show_custom_modules"];
            }
            set
            {
                this["show_custom_modules"] = value;
            }
        }
        [DefaultSettingValue("<?xml version=\"1.0\" encoding=\"utf-16\"?>\r\n<ArrayOfString xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">\r\n  <string>None|None</string>\r\n</ArrayOfString>"), UserScopedSetting, DebuggerNonUserCode]
        public StringCollection custom_modules
        {
            get
            {
                return (StringCollection)this["custom_modules"];
            }
            set
            {
                this["custom_modules"] = value;
            }
        }
        [DefaultSettingValue("1000"), DebuggerNonUserCode, UserScopedSetting]
        public int sync_max_records
        {
            get
            {
                return (int)this["sync_max_records"];
            }
            set
            {
                this["sync_max_records"] = value;
            }
        }
        [DebuggerNonUserCode, UserScopedSetting, DefaultSettingValue("1,2,3")]
        public string selected_search_modules
        {
            get
            {
                return (string)this["selected_search_modules"];
            }
            set
            {
                this["selected_search_modules"] = value;
            }
        }
        [UserScopedSetting, DebuggerNonUserCode, DefaultSettingValue("False")]
        public bool participate_in_ceip
        {
            get
            {
                return (bool)this["participate_in_ceip"];
            }
            set
            {
                this["participate_in_ceip"] = value;
            }
        }
        [DefaultSettingValue("True"), DebuggerNonUserCode, UserScopedSetting]
        public bool populate_context_lookup_list
        {
            get
            {
                return (bool)this["populate_context_lookup_list"];
            }
            set
            {
                this["populate_context_lookup_list"] = value;
            }
        }
        [DefaultSettingValue("True"), DebuggerNonUserCode, UserScopedSetting]
        public bool auto_archive
        {
            get
            {
                return (bool)this["auto_archive"];
            }
            set
            {
                this["auto_archive"] = value;
            }
        }
        [DefaultSettingValue(""), DebuggerNonUserCode, UserScopedSetting]
        public System.Collections.Generic.List<string> auto_archive_folders
        {
            get
            {
                return (System.Collections.Generic.List<string>)this["auto_archive_folders"];               
            }
            set
            {
                this["auto_archive_folders"] = value;
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

        [DefaultSettingValue("False"), DebuggerNonUserCode, UserScopedSetting]
        public bool AttachmentsChecked
        {
            get
            {
                return (bool)this["AttachmentsChecked"];
            }
            set
            {
                this["AttachmentsChecked"] = value;
            }
        }

        [UserScopedSetting, DebuggerNonUserCode]
        public StringCollection opportunity_dropdown_salestage
        {
            get
            {
                return (StringCollection)this["opportunity_dropdown_salestage"];
            }
            set
            {
                this["opportunity_dropdown_salestage"] = value;
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
        public static Hashtable accountEntrys = new Hashtable();
        public static AutoCompleteStringCollection accounts;
    }
}
