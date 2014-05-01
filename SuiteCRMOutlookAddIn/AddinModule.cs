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
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Windows.Forms;
using AddinExpress.MSO;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.VisualBasic;
using System.Collections;
using System.Collections.Generic;
using SuiteCRMClient;

namespace SuiteCRMOutlookAddIn
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("A232451B-DD35-4A43-9B88-C23226D039B7"), ProgId("SuiteCRMOutlookAddIn.AddinModule")]
    public class AddinModule : AddinExpress.MSO.ADXAddinModule
    {
        public AddinModule()
        {
            Application.EnableVisualStyles();
            InitializeComponent();
            // Please add any initialization code to the AddinInitialize event handler
        }
        public clsSettings settings;
        private ADXRibbonTab SuiteCRMTab;
        private ADXRibbonGroup adxRibbonGroup1;
        private ADXOutlookAppEvents adxOutlookEvents;
        private ADXContextMenu adxContextMenu1;
        private ADXCommandBarButton cbbSugarCRMArcive;
        private ImageList AllImages;
        private ADXRibbonContextMenu rcmMainMenu;
        private ADXRibbonButton rnSugarCRMArchive;
        private ADXOlExplorerMainMenu adxOlExplorerMainMenu1;
        private ADXCommandBarPopup adxCommandBarPopup1;
        private ADXCommandBarButton acbbArchive;
        private ADXCommandBarButton acbbSettings;
        private ADXRibbonTab SuiteCRMComposeTab;
        private ADXRibbonGroup adxRibbonGroup3;
        private ADXRibbonButton arbAddressBook;
        private ADXRibbonGroup adxRibbonGroup2;
        private ADXRibbonButton adxRibbonButton1;
        private ADXRibbonButton adxRibbonButton2;
        private ADXRibbonButton adxRibbonButton3;
        private ADXRibbonButton arbComposeSettings;
        private ADXRibbonButton adxRibbonButtonArchive;
        private ADXRibbonButton adxRibbonButtonSettings;

        #region Component Designer generated code
        /// <summary>
        /// Required by designer
        /// </summary>
        private System.ComponentModel.IContainer components;

        /// <summary>
        /// Required by designer support - do not modify
        /// the following method
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddinModule));
            this.SuiteCRMTab = new AddinExpress.MSO.ADXRibbonTab(this.components);
            this.adxRibbonGroup1 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.adxRibbonButtonArchive = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.AllImages = new System.Windows.Forms.ImageList(this.components);
            this.adxRibbonButtonSettings = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxOutlookEvents = new AddinExpress.MSO.ADXOutlookAppEvents(this.components);
            this.adxContextMenu1 = new AddinExpress.MSO.ADXContextMenu(this.components);
            this.cbbSugarCRMArcive = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.rcmMainMenu = new AddinExpress.MSO.ADXRibbonContextMenu(this.components);
            this.rnSugarCRMArchive = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxOlExplorerMainMenu1 = new AddinExpress.MSO.ADXOlExplorerMainMenu(this.components);
            this.adxCommandBarPopup1 = new AddinExpress.MSO.ADXCommandBarPopup(this.components);
            this.acbbArchive = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.acbbSettings = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.SuiteCRMComposeTab = new AddinExpress.MSO.ADXRibbonTab(this.components);
            this.adxRibbonGroup3 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.arbAddressBook = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.arbComposeSettings = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonGroup2 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.adxRibbonButton1 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonButton2 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonButton3 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            // 
            // SuiteCRMTab
            // 
            this.SuiteCRMTab.Caption = "SuiteCRM";
            this.SuiteCRMTab.Controls.Add(this.adxRibbonGroup1);
            this.SuiteCRMTab.Id = "adxRibbonTab_a15b396c24834b54a6436f07eff369d0";
            this.SuiteCRMTab.Ribbons = ((AddinExpress.MSO.ADXRibbons)((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead | AddinExpress.MSO.ADXRibbons.msrOutlookExplorer)));
            // 
            // adxRibbonGroup1
            // 
            this.adxRibbonGroup1.Caption = "SuiteCRM";
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonArchive);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonSettings);
            this.adxRibbonGroup1.Id = "adxRibbonGroup_df8c10d835c240da8266cc8250e77d50";
            this.adxRibbonGroup1.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup1.Ribbons = ((AddinExpress.MSO.ADXRibbons)((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead | AddinExpress.MSO.ADXRibbons.msrOutlookExplorer)));
            // 
            // adxRibbonButtonArchive
            // 
            this.adxRibbonButtonArchive.Caption = "Archive";
            this.adxRibbonButtonArchive.Id = "adxRibbonButton_bbac3bbfbdb644ad96cef7d78970a853";
            this.adxRibbonButtonArchive.Image = 3;
            this.adxRibbonButtonArchive.ImageList = this.AllImages;
            this.adxRibbonButtonArchive.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonArchive.Ribbons = ((AddinExpress.MSO.ADXRibbons)((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead | AddinExpress.MSO.ADXRibbons.msrOutlookExplorer)));
            this.adxRibbonButtonArchive.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonArchive.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButtonArchive_OnClick);
            // 
            // AllImages
            // 
            this.AllImages.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("AllImages.ImageStream")));
            this.AllImages.TransparentColor = System.Drawing.Color.Transparent;
            this.AllImages.Images.SetKeyName(0, "Settings.png");
            this.AllImages.Images.SetKeyName(1, "SugarCRM.png");
            this.AllImages.Images.SetKeyName(2, "Contacts.jpg");
            this.AllImages.Images.SetKeyName(3, "SuiteCRM.jpg");
            this.AllImages.Images.SetKeyName(4, "AddOpportunities.jpg");
            // 
            // adxRibbonButtonSettings
            // 
            this.adxRibbonButtonSettings.Caption = "Settings";
            this.adxRibbonButtonSettings.Id = "adxRibbonButton_25265e2fb82c4a488c51f4b60dec884c";
            this.adxRibbonButtonSettings.Image = 0;
            this.adxRibbonButtonSettings.ImageList = this.AllImages;
            this.adxRibbonButtonSettings.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonSettings.Ribbons = ((AddinExpress.MSO.ADXRibbons)((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead | AddinExpress.MSO.ADXRibbons.msrOutlookExplorer)));
            this.adxRibbonButtonSettings.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonSettings.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButtonSettings_OnClick);
            // 
            // adxOutlookEvents
            // 
            this.adxOutlookEvents.ItemSend += new AddinExpress.MSO.ADXOlItemSend_EventHandler(this.adxOutlookEvents_ItemSend);
            this.adxOutlookEvents.Startup += new System.EventHandler(this.adxOutlookEvents_Startup);
            this.adxOutlookEvents.Quit += new System.EventHandler(this.adxOutlookEvents_Quit);
            this.adxOutlookEvents.ItemContextMenuDisplay += new AddinExpress.MSO.ADXOlContextMenu_EventHandler(this.adxOutlookEvents_ItemContextMenuDisplay);
            // 
            // adxContextMenu1
            // 
            this.adxContextMenu1.CommandBarName = "Context Menu";
            this.adxContextMenu1.CommandBarTag = "ba4409b3-812a-406e-8c08-ad8af76f14ea";
            this.adxContextMenu1.Controls.Add(this.cbbSugarCRMArcive);
            this.adxContextMenu1.SupportedApp = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;
            this.adxContextMenu1.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;
            this.adxContextMenu1.Temporary = true;
            this.adxContextMenu1.UpdateCounter = 3;
            // 
            // cbbSugarCRMArcive
            // 
            this.cbbSugarCRMArcive.Caption = "SuiteCRM Archive";
            this.cbbSugarCRMArcive.ControlTag = "f0304722-dffc-470c-9005-f5ff87a778bf";
            this.cbbSugarCRMArcive.DescriptionText = "SuiteCRM Archive";
            this.cbbSugarCRMArcive.Image = 3;
            this.cbbSugarCRMArcive.ImageList = this.AllImages;
            this.cbbSugarCRMArcive.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.cbbSugarCRMArcive.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.cbbSugarCRMArcive.Temporary = true;
            this.cbbSugarCRMArcive.TooltipText = "SugarCRM Arcive";
            this.cbbSugarCRMArcive.UpdateCounter = 20;
            this.cbbSugarCRMArcive.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.cbbSugarCRMArcive_Click);
            // 
            // rcmMainMenu
            // 
            this.rcmMainMenu.ContextMenuNames.AddRange(new string[] {
            "Outlook.Explorer.ContextMenuMailItem",
            "Outlook.Explorer.ContextMenuMultipleItems"});
            this.rcmMainMenu.Controls.Add(this.rnSugarCRMArchive);
            this.rcmMainMenu.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
            // 
            // rnSugarCRMArchive
            // 
            this.rnSugarCRMArchive.Caption = "SuiteCRM Archive";
            this.rnSugarCRMArchive.Description = "SuiteCRM Archive";
            this.rnSugarCRMArchive.Id = "adxRibbonButton_fa88689ac9634d5ca68fa6eb600e5366";
            this.rnSugarCRMArchive.Image = 3;
            this.rnSugarCRMArchive.ImageList = this.AllImages;
            this.rnSugarCRMArchive.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.rnSugarCRMArchive.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
            this.rnSugarCRMArchive.ScreenTip = "SugarCRM Arcive";
            this.rnSugarCRMArchive.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.rnSugarCRMArchive.SuperTip = "SugarCRM Arcive";
            this.rnSugarCRMArchive.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.rnSugarCRMArcive_OnClick);
            // 
            // adxOlExplorerMainMenu1
            // 
            this.adxOlExplorerMainMenu1.CommandBarName = "Menu Bar";
            this.adxOlExplorerMainMenu1.CommandBarTag = "5090e07a-beaa-4da1-9111-3069aea7a855";
            this.adxOlExplorerMainMenu1.Controls.Add(this.adxCommandBarPopup1);
            this.adxOlExplorerMainMenu1.Temporary = true;
            this.adxOlExplorerMainMenu1.UpdateCounter = 4;
            this.adxOlExplorerMainMenu1.UseForRibbon = true;
            // 
            // adxCommandBarPopup1
            // 
            this.adxCommandBarPopup1.Caption = "SuiteCRM";
            this.adxCommandBarPopup1.Controls.Add(this.acbbArchive);
            this.adxCommandBarPopup1.Controls.Add(this.acbbSettings);
            this.adxCommandBarPopup1.ControlTag = "3e33f2a8-ec42-496d-b676-8b93a6b7cda7";
            this.adxCommandBarPopup1.DescriptionText = "SuiteCRM";
            this.adxCommandBarPopup1.Temporary = true;
            this.adxCommandBarPopup1.TooltipText = "SuiteCRM";
            this.adxCommandBarPopup1.UpdateCounter = 3;
            // 
            // acbbArchive
            // 
            this.acbbArchive.Caption = "Archive";
            this.acbbArchive.ControlTag = "b21305bd-4bea-42f5-89f8-a8da78898f48";
            this.acbbArchive.DescriptionText = "Archive";
            this.acbbArchive.Image = 1;
            this.acbbArchive.ImageList = this.AllImages;
            this.acbbArchive.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.acbbArchive.Temporary = true;
            this.acbbArchive.TooltipText = "Archive";
            this.acbbArchive.UpdateCounter = 9;
            this.acbbArchive.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.acbbArchive_Click);
            // 
            // acbbSettings
            // 
            this.acbbSettings.Caption = "Settings";
            this.acbbSettings.ControlTag = "97901bd1-5a4e-402c-9038-9a70db1c9e9b";
            this.acbbSettings.DescriptionText = "Settings";
            this.acbbSettings.Image = 0;
            this.acbbSettings.ImageList = this.AllImages;
            this.acbbSettings.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.acbbSettings.Temporary = true;
            this.acbbSettings.TooltipText = "Settings";
            this.acbbSettings.UpdateCounter = 9;
            this.acbbSettings.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.acbbSettings_Click);
            // 
            // SuiteCRMComposeTab
            // 
            this.SuiteCRMComposeTab.Caption = "SuiteCRM";
            this.SuiteCRMComposeTab.Controls.Add(this.adxRibbonGroup3);
            this.SuiteCRMComposeTab.Id = "adxRibbonTab_8b4fb968dc014a64aae891c9292452fa";
            this.SuiteCRMComposeTab.IdMso = "TabNewMailMessage";
            this.SuiteCRMComposeTab.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose;
            // 
            // adxRibbonGroup3
            // 
            this.adxRibbonGroup3.Caption = "SuiteCRM";
            this.adxRibbonGroup3.Controls.Add(this.arbAddressBook);
            this.adxRibbonGroup3.Controls.Add(this.arbComposeSettings);
            this.adxRibbonGroup3.Id = "adxRibbonGroup_d9f85d68f4f64ffe95b1b4ef12b154cb";
            this.adxRibbonGroup3.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup3.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose;
            // 
            // arbAddressBook
            // 
            this.arbAddressBook.Caption = "Address Book";
            this.arbAddressBook.Description = "SuiteCRM Address Book";
            this.arbAddressBook.Id = "adxRibbonButton_9b3442871bd3407080c0f608ae5fa231";
            this.arbAddressBook.Image = 3;
            this.arbAddressBook.ImageList = this.AllImages;
            this.arbAddressBook.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.arbAddressBook.KeyTip = "Sui";
            this.arbAddressBook.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose;
            this.arbAddressBook.ScreenTip = "SuiteCRM Address Book";
            this.arbAddressBook.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.arbAddressBook.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.arbAddressBook_OnClick);
            // 
            // arbComposeSettings
            // 
            this.arbComposeSettings.Caption = "Settings";
            this.arbComposeSettings.Description = "Settings";
            this.arbComposeSettings.Id = "adxRibbonButton_88a3d408d17d472e8560cf0fe9d5899b";
            this.arbComposeSettings.Image = 0;
            this.arbComposeSettings.ImageList = this.AllImages;
            this.arbComposeSettings.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.arbComposeSettings.KeyTip = "Set";
            this.arbComposeSettings.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookMailCompose;
            this.arbComposeSettings.ScreenTip = "Settings";
            this.arbComposeSettings.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.arbComposeSettings.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.arbComposeSettings_OnClick);
            // 
            // adxRibbonGroup2
            // 
            this.adxRibbonGroup2.Caption = "SuiteCRM";
            this.adxRibbonGroup2.Controls.Add(this.adxRibbonButton1);
            this.adxRibbonGroup2.Controls.Add(this.adxRibbonButton2);
            this.adxRibbonGroup2.Controls.Add(this.adxRibbonButton3);
            this.adxRibbonGroup2.Id = "adxRibbonGroup_df8c10d835c240da8266cc8250e77d50";
            this.adxRibbonGroup2.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup2.Ribbons = ((AddinExpress.MSO.ADXRibbons)((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead | AddinExpress.MSO.ADXRibbons.msrOutlookExplorer)));
            // 
            // adxRibbonButton1
            // 
            this.adxRibbonButton1.Caption = "Archive";
            this.adxRibbonButton1.Id = "adxRibbonButton_080b5520a8704cdab2255540f605c574";
            this.adxRibbonButton1.ImageList = this.AllImages;
            this.adxRibbonButton1.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton1.Ribbons = ((AddinExpress.MSO.ADXRibbons)((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead | AddinExpress.MSO.ADXRibbons.msrOutlookExplorer)));
            this.adxRibbonButton1.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            // 
            // adxRibbonButton2
            // 
            this.adxRibbonButton2.Caption = "Add Contact";
            this.adxRibbonButton2.Id = "adxRibbonButton_d8059f77747e48aa8724c887cba44140";
            this.adxRibbonButton2.Image = 2;
            this.adxRibbonButton2.ImageList = this.AllImages;
            this.adxRibbonButton2.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton2.Ribbons = ((AddinExpress.MSO.ADXRibbons)((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead | AddinExpress.MSO.ADXRibbons.msrOutlookExplorer)));
            this.adxRibbonButton2.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButton2.Visible = false;
            // 
            // adxRibbonButton3
            // 
            this.adxRibbonButton3.Caption = "Settings";
            this.adxRibbonButton3.Id = "adxRibbonButton_3e2e0c3ae23b488b92c642975a6ebf58";
            this.adxRibbonButton3.Image = 0;
            this.adxRibbonButton3.ImageList = this.AllImages;
            this.adxRibbonButton3.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton3.Ribbons = ((AddinExpress.MSO.ADXRibbons)((AddinExpress.MSO.ADXRibbons.msrOutlookMailRead | AddinExpress.MSO.ADXRibbons.msrOutlookExplorer)));
            this.adxRibbonButton3.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            // 
            // AddinModule
            // 
            this.AddinName = "SuiteCRMOutlookAddIn";
            this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;

        }
        #endregion

        #region Add-in Express automatic code

        // Required by Add-in Express - do not modify
        // the methods within this region

        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }

        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
        }

        [ComUnregisterFunctionAttribute]
        public static void AddinUnregister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXUnregister(t);
        }

        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

        public static new AddinModule CurrentInstance
        {
            get
            {
                return AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule;
            }
        }

        public Outlook._Application OutlookApp
        {
            get
            {
                return (HostApplication as Outlook._Application);
            }
        }

        public SuiteCRMClient.clsUsersession SugarCRMUserSession;

        private void ManualArchive()
        {
            SugarCRMAuthenticate();
            frmArchive objForm = new frmArchive();
            objForm.ShowDialog();
        }

        private void ArchiveEmail(Outlook.MailItem objMail, int intArchiveType, string strExcludedEmails="")
        {
            SuiteCRMClient.clsEmailArchive objEmail = new SuiteCRMClient.clsEmailArchive();
            objEmail.From = objMail.SenderEmailAddress;
            objEmail.To = "";
            foreach (Outlook.Recipient objRecepient in objMail.Recipients)
            {
                if (objEmail.To == "")
                    objEmail.To = objRecepient.Address;
                else
                    objEmail.To += ";" + objRecepient.Address;
            }
            objEmail.Subject = objMail.Subject;
            objEmail.Body = objMail.Body;
            objEmail.HTMLBody = objMail.HTMLBody;
            objEmail.ArchiveType = intArchiveType;
            foreach (Outlook.Attachment objMailAttachments in objMail.Attachments)
            {
                objEmail.Attachments.Add(new SuiteCRMClient.clsEmailAttachments { DisplayName = objMailAttachments.DisplayName, FileContentInBase64String = Base64Encode(objMailAttachments, objMail) });
            }

            System.Threading.Thread objThread = new System.Threading.Thread(() => ArchiveEmailThread(objEmail, intArchiveType, strExcludedEmails));
            objThread.Start();
        }

        public static byte[] Base64Encode(Outlook.Attachment objMailAttachment, Outlook.MailItem objMail)
        {
            byte[] strRet = null;
            if (objMailAttachment != null)
            {
                if (System.IO.Directory.Exists(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath") == false)
                {
                    System.IO.Directory.CreateDirectory(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath");
                }
                try
                {
                    objMailAttachment.SaveAsFile(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath\\" + objMailAttachment.FileName);
                    strRet = System.IO.File.ReadAllBytes(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath\\" + objMailAttachment.FileName);
                }
                catch (COMException ex)
                {
                    try
                    {
                        clsSuiteCRMHelper.LoadLogFileLocation();
                        clsSuiteCRMHelper.AddLogLine("------------------" + System.DateTime.Now.ToString() + "-----------------");
                        clsSuiteCRMHelper.AddLogLine("AddInModule.Base64Encode method COM Exception:");
                        clsSuiteCRMHelper.AddLogLine("Message:" + ex.Message);
                        clsSuiteCRMHelper.AddLogLine("Source:" + ex.Source);
                        clsSuiteCRMHelper.AddLogLine("StackTrace:" + ex.StackTrace);
                        clsSuiteCRMHelper.AddLogLine("Data:" + ex.Data.ToString());
                        clsSuiteCRMHelper.AddLogLine("HResult:" + ex.HResult.ToString());
                        clsSuiteCRMHelper.AddLogLine("Inputs:");
                        clsSuiteCRMHelper.AddLogLine("Data:" + objMailAttachment.DisplayName);
                        clsSuiteCRMHelper.AddLogLine("-------------------------------------------------------------------------");
                        clsSuiteCRMHelper.log.Close();
                        ex.Data.Clear();
                        string strName = Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath\\" + DateTime.Now.ToString("MMddyyyyHHmmssfff") + ".html";
                        objMail.SaveAs(strName, Microsoft.Office.Interop.Outlook.OlSaveAsType.olHTML);
                        foreach (string strFileName in System.IO.Directory.GetFiles(strName.Replace(".html", "_files")))
                        {
                            if (strFileName.EndsWith("\\" + objMailAttachment.DisplayName))
                            {
                                strRet = System.IO.File.ReadAllBytes(strFileName);
                                break;
                            }
                        }
                    }
                    catch (Exception ex1)
                    {
                        clsSuiteCRMHelper.LoadLogFileLocation();
                        clsSuiteCRMHelper.AddLogLine("------------------" + System.DateTime.Now.ToString() + "-----------------");
                        clsSuiteCRMHelper.AddLogLine("AddInModule.Base64Encode method General Exception:");
                        clsSuiteCRMHelper.AddLogLine("Message:" + ex.Message);
                        clsSuiteCRMHelper.AddLogLine("Source:" + ex.Source);
                        clsSuiteCRMHelper.AddLogLine("StackTrace:" + ex.StackTrace);
                        clsSuiteCRMHelper.AddLogLine("Data:" + ex.Data.ToString());
                        clsSuiteCRMHelper.AddLogLine("HResult:" + ex.HResult.ToString());
                        clsSuiteCRMHelper.AddLogLine("Inputs:");
                        clsSuiteCRMHelper.AddLogLine("Data:" + objMailAttachment.DisplayName);
                        clsSuiteCRMHelper.AddLogLine("-------------------------------------------------------------------------");
                        clsSuiteCRMHelper.log.Close();
                        ex1.Data.Clear();
                    }
                }
                finally
                {
                    if (System.IO.Directory.Exists(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath") == true)
                    {
                        System.IO.Directory.Delete(Environment.SpecialFolder.MyDocuments.ToString() + "\\SuiteCRMTempAttachmentPath", true);
                    }
                }
            }

            return strRet;
        }


        private void ArchiveEmailThread(SuiteCRMClient.clsEmailArchive objEmail, int intArchiveType, string strExcludedEmails = "")
        {
            SugarCRMAuthenticate();
            if (SugarCRMUserSession != null)
            {
                while (SugarCRMUserSession.AwaitingAuthentication == true)
                {
                    System.Threading.Thread.Sleep(1000);
                }
            }
            SugarCRMAuthenticate();
            objEmail.SugarCRMUserSession = SugarCRMUserSession;
            objEmail.Save(strExcludedEmails);

        }

        public void SugarCRMAuthenticate()
        {
            if (SugarCRMUserSession == null)
            {
                Authenticate();
            }
            else
            {
                if (SugarCRMUserSession.id == "")
                    Authenticate();
            }

        }

        public void Authenticate()
        {
            string strURL = settings.host;
            if (strURL != "")
            {
                string strUsername = settings.username;
                string strPassword = settings.password;
                SugarCRMUserSession = new SuiteCRMClient.clsUsersession(strURL, strUsername, strPassword);
                SugarCRMUserSession.AwaitingAuthentication = true;
                try
                {
                    SugarCRMUserSession.Login();
                    if (SugarCRMUserSession.id != "")
                        return;
                }
                catch (Exception ex)
                {
                    ex.Data.Clear();
                }
            }
            frmSettings objSettings = new frmSettings();
            objSettings.ShowDialog();
            SugarCRMUserSession.AwaitingAuthentication = false;
        }
      

        private void adxOutlookEvents_Quit(object sender, EventArgs e)
        {
            try
            {
                if (SugarCRMUserSession != null)
                    SugarCRMUserSession.LogOut();
            }
            catch (Exception ex)
            {
                ex.Data.Clear();
            }
        }


        List<Outlook.Folder> lstOutlookFolders;
        private void GetMailFolders(Outlook.Folders objInpFolders)
        {
            foreach (Outlook.Folder objFolder in objInpFolders)
            {
                if (objFolder.Folders.Count > 0)
                {
                    lstOutlookFolders.Add(objFolder);
                    GetMailFolders(objFolder.Folders);
                }
                else
                    lstOutlookFolders.Add(objFolder);
            }
        }

        private void ArchiveFolderItems(Outlook.Folder objFolder, DateTime? dtAutoArchiveFrom = null)
        {
            Outlook.Items UnReads;
            if (dtAutoArchiveFrom== null)
                UnReads = objFolder.Items.Restrict("[Unread]=true");
            else
                UnReads = objFolder.Items.Restrict("[ReceivedTime] >= '" + ((DateTime)dtAutoArchiveFrom).AddDays(-1).ToString("yyyy-MM-dd HH:mm") + "'");

            for (int intItr = 1; intItr <= UnReads.Count; intItr++)
            {
                if (UnReads[intItr] is Outlook.MailItem)
                {
                    Outlook.MailItem objMail = (Outlook.MailItem)UnReads[intItr];                                     

                        if (objMail.UserProperties["SuiteCRM"] == null)
                        {
                            ArchiveEmail(objMail, 2, this.settings.ExcludedEmails);
                            objMail.UserProperties.Add("SuiteCRM", Outlook.OlUserPropertyType.olText, true, Outlook.OlUserPropertyType.olText);
                            objMail.UserProperties["SuiteCRM"].Value = "True";
                            objMail.Save();
                        }
                        else
                            break;                    
                }
            }
        }

        public void ProcessMails(DateTime? dtAutoArchiveFrom = null)
        {
            if (settings.auto_archive == false)
                return;
            System.Threading.Thread.Sleep(5000);
            while (1 == 1)
            {
                try
                {
                    lstOutlookFolders = new List<Outlook.Folder>();
                    GetMailFolders(OutlookApp.Session.Folders);
                    if (lstOutlookFolders != null)
                    {
                        foreach (Outlook.Folder objFolder in lstOutlookFolders)
                        {
                            if (settings.auto_archive_folders == null)
                                ArchiveFolderItems(objFolder, dtAutoArchiveFrom);
                            else if (settings.auto_archive_folders.Count == 0)
                                ArchiveFolderItems(objFolder, dtAutoArchiveFrom);
                            else
                            {
                                if (settings.auto_archive_folders.Contains(objFolder.EntryID))
                                {
                                    ArchiveFolderItems(objFolder, dtAutoArchiveFrom);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    ex.Data.Clear();
                }
                if (dtAutoArchiveFrom != null)
                    break;

                System.Threading.Thread.Sleep(5000);
            }
        }


        //Commandbars Right Click (2003-2007)
        private void cbbSugarCRMArcive_Click(object sender)
        {
            ManualArchive();
        }

        //Ribbon Type Outlook Right Click (2003-2007)
        private void rnSugarCRMArcive_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            ManualArchive();
        }

        Outlook.Selection m_selection;
        private void adxOutlookEvents_ItemContextMenuDisplay(object sender, object commandBar, object target)
        {
            m_selection = target as Outlook.Selection;
        }

        private void acbbSettings_Click(object sender)
        {
            frmSettings objacbbSettings = new frmSettings();
            objacbbSettings.ShowDialog();
        }

        private void acbbArchive_Click(object sender)
        {
            ManualArchive();
        }

        private void adxOutlookEvents_Startup(object sender, EventArgs e)
        {
            SuiteCRMClient.clsSuiteCRMHelper.InstallationPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SuiteCRMOutlookAddIn";
            this.settings = new clsSettings();
            System.Threading.Thread objThread = new System.Threading.Thread(() => ProcessMails());
            objThread.Start();
        }

        private void adxOutlookEvents_ItemSend(object sender, ADXOlItemSendEventArgs e)
        {
                if (e.Item is Outlook.MailItem)
                {
                    Outlook.MailItem objMail = (Outlook.MailItem)e.Item;
                    if (objMail.UserProperties["SuiteCRM"] == null)
                    {
                        ArchiveEmail(objMail, 1, this.settings.ExcludedEmails);
                        objMail.UserProperties.Add("SuiteCRM", Outlook.OlUserPropertyType.olText, true, Outlook.OlUserPropertyType.olText);
                        objMail.UserProperties["SuiteCRM"].Value = "True";
                        objMail.Save();
                    }
                }
        }

        private void arbAddressBook_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            frmAddressBook objForm = new frmAddressBook();
            objForm.Show();
        }

        private void arbComposeSettings_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            frmSettings objacbbSettings = new frmSettings();
            objacbbSettings.ShowDialog();
        }

       
        private void adxRibbonButtonArchive_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            ManualArchive();
        }

        private void adxRibbonButtonSettings_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            frmSettings objSettings = new frmSettings();
            objSettings.ShowDialog();
        }
    }
}

