namespace XLEdge
{
    partial class AddinModule
    {
        /// <summary>
        /// Required by designer
        /// </summary>
        private System.ComponentModel.IContainer components;
 
        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code
        /// <summary>
        /// Required by designer support - do not modify
        /// the following method
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddinModule));
            this.TabXLEdge = new AddinExpress.MSO.ADXRibbonTab(this.components);
            this.adxRibbonGroup1 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.RibEdgeLogin = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.ImageList_32X32 = new System.Windows.Forms.ImageList(this.components);
            this.RibEdgeLogout = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.RibLoginURL = new AddinExpress.MSO.ADXRibbonLabel(this.components);
            this.RibEdgeDialogBoxLauncher = new AddinExpress.MSO.ADXRibbonDialogBoxLauncher(this.components);
            this.adxRibbonGroup2 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.RibEdgeRefresh = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.RibEdgeRefreshAll = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonGroup3 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.RibEdgeParamRefresh = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.RibEdgeParamRefreshBook = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonGroup4 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.RibEdgeShowHide = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonGroup5 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.RibEdgeDebug = new AddinExpress.MSO.ADXRibbonCheckBox(this.components);
            this.RibEdgeIncludeOutputData = new AddinExpress.MSO.ADXRibbonCheckBox(this.components);
            this.adxRibbonGroup6 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.RibEdgeOptions = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.RibControlSheet = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonGroup7 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.RibEdgeAbout = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonGroup8 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.RibEdgeHelp = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonGroup9 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.RibSheetLabel = new AddinExpress.MSO.ADXRibbonLabel(this.components);
            this.adxExcelTaskPanesManager1 = new AddinExpress.XL.ADXExcelTaskPanesManager(this.components);
            this.adxExcelTaskPanesCollectionItem1 = new AddinExpress.XL.ADXExcelTaskPanesCollectionItem(this.components);
            this.adxExcelAppEvents1 = new AddinExpress.MSO.ADXExcelAppEvents(this.components);
            // 
            // TabXLEdge
            // 
            this.TabXLEdge.Caption = "Orbit XLEdge";
            this.TabXLEdge.Controls.Add(this.adxRibbonGroup1);
            this.TabXLEdge.Controls.Add(this.adxRibbonGroup2);
            this.TabXLEdge.Controls.Add(this.adxRibbonGroup3);
            this.TabXLEdge.Controls.Add(this.adxRibbonGroup4);
            this.TabXLEdge.Controls.Add(this.adxRibbonGroup5);
            this.TabXLEdge.Controls.Add(this.adxRibbonGroup6);
            this.TabXLEdge.Controls.Add(this.adxRibbonGroup7);
            this.TabXLEdge.Controls.Add(this.adxRibbonGroup8);
            this.TabXLEdge.Controls.Add(this.adxRibbonGroup9);
            this.TabXLEdge.Id = "adxRibbonTab_d1d317b20611427bbcbf7c517182c3bd";
            this.TabXLEdge.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonGroup1
            // 
            this.adxRibbonGroup1.Caption = "Login";
            this.adxRibbonGroup1.Controls.Add(this.RibEdgeLogin);
            this.adxRibbonGroup1.Controls.Add(this.RibEdgeLogout);
            this.adxRibbonGroup1.Controls.Add(this.RibLoginURL);
            this.adxRibbonGroup1.Controls.Add(this.RibEdgeDialogBoxLauncher);
            this.adxRibbonGroup1.Id = "adxRibbonGroup_e94df3d8ba104947b30c6d3bc39c8901";
            this.adxRibbonGroup1.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // RibEdgeLogin
            // 
            this.RibEdgeLogin.Caption = "Login";
            this.RibEdgeLogin.Id = "adxRibbonButton_7a0d71e54106472d9910d134448b1f26";
            this.RibEdgeLogin.Image = 0;
            this.RibEdgeLogin.ImageList = this.ImageList_32X32;
            this.RibEdgeLogin.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.RibEdgeLogin.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibEdgeLogin.ScreenTip = "Login";
            this.RibEdgeLogin.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.RibEdgeLogin.SuperTip = "Click to login";
            this.RibEdgeLogin.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibLogin_OnClick);
            // 
            // ImageList_32X32
            // 
            this.ImageList_32X32.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("ImageList_32X32.ImageStream")));
            this.ImageList_32X32.TransparentColor = System.Drawing.Color.Transparent;
            this.ImageList_32X32.Images.SetKeyName(0, "login.png");
            this.ImageList_32X32.Images.SetKeyName(1, "logout.png");
            this.ImageList_32X32.Images.SetKeyName(2, "worksheet-refersh_1.png");
            this.ImageList_32X32.Images.SetKeyName(3, "workbook-refersh_1.png");
            this.ImageList_32X32.Images.SetKeyName(4, "worksheet__Rerun.png");
            this.ImageList_32X32.Images.SetKeyName(5, "workbook_Rerun.png");
            this.ImageList_32X32.Images.SetKeyName(6, "reports-1-01.png");
            this.ImageList_32X32.Images.SetKeyName(7, "options-1-01 (2).png");
            this.ImageList_32X32.Images.SetKeyName(8, "create-control-panel-01.png");
            this.ImageList_32X32.Images.SetKeyName(9, "About.png");
            this.ImageList_32X32.Images.SetKeyName(10, "Help.png");
            // 
            // RibEdgeLogout
            // 
            this.RibEdgeLogout.Caption = "Logout";
            this.RibEdgeLogout.Id = "adxRibbonButton_eda1721fd9ad412f96f972e6bfc49612";
            this.RibEdgeLogout.Image = 1;
            this.RibEdgeLogout.ImageList = this.ImageList_32X32;
            this.RibEdgeLogout.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.RibEdgeLogout.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibEdgeLogout.ScreenTip = "Logout";
            this.RibEdgeLogout.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.RibEdgeLogout.SuperTip = "Click to logout";
            this.RibEdgeLogout.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibEdgeLogout_OnClick);
            this.RibEdgeLogout.PropertyChanging += new AddinExpress.MSO.ADXRibbonPropertyChanging_EventHandler(this.RibEdgeLogout_PropertyChanging);
            // 
            // RibLoginURL
            // 
            this.RibLoginURL.Caption = " ";
            this.RibLoginURL.Id = "adxRibbonLabel_2304ea98266b40628c3467dd8928f38c";
            this.RibLoginURL.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibLoginURL.Visible = false;
            // 
            // RibEdgeDialogBoxLauncher
            // 
            this.RibEdgeDialogBoxLauncher.Id = "adxRibbonDialogBoxLauncher_f36cdeedab704ec48030a4c60c1fa923";
            this.RibEdgeDialogBoxLauncher.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibEdgeDialogBoxLauncher.ScreenTip = "Logged-in Details";
            this.RibEdgeDialogBoxLauncher.SuperTip = "User and instance logged in details.";
            this.RibEdgeDialogBoxLauncher.OnAction += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibEdgeDialogBoxLauncher_OnAction);
            // 
            // adxRibbonGroup2
            // 
            this.adxRibbonGroup2.Caption = "Refresh";
            this.adxRibbonGroup2.Controls.Add(this.RibEdgeRefresh);
            this.adxRibbonGroup2.Controls.Add(this.RibEdgeRefreshAll);
            this.adxRibbonGroup2.Id = "adxRibbonGroup_c8b5a977cf774ddf9e33f8a9c1783c19";
            this.adxRibbonGroup2.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup2.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // RibEdgeRefresh
            // 
            this.RibEdgeRefresh.Caption = "Sheet";
            this.RibEdgeRefresh.Id = "adxRibbonButton_32c5a1c8b3cd45bc9053375d9fb532a5";
            this.RibEdgeRefresh.Image = 2;
            this.RibEdgeRefresh.ImageList = this.ImageList_32X32;
            this.RibEdgeRefresh.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.RibEdgeRefresh.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibEdgeRefresh.ScreenTip = "Worksheet";
            this.RibEdgeRefresh.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.RibEdgeRefresh.SuperTip = resources.GetString("RibEdgeRefresh.SuperTip");
            this.RibEdgeRefresh.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibEdgeRefresh_OnClick);
            // 
            // RibEdgeRefreshAll
            // 
            this.RibEdgeRefreshAll.Caption = "Book";
            this.RibEdgeRefreshAll.Id = "adxRibbonButton_6379a68429884c129beafd740f9662b9";
            this.RibEdgeRefreshAll.Image = 3;
            this.RibEdgeRefreshAll.ImageList = this.ImageList_32X32;
            this.RibEdgeRefreshAll.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.RibEdgeRefreshAll.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibEdgeRefreshAll.ScreenTip = "Workbook";
            this.RibEdgeRefreshAll.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.RibEdgeRefreshAll.SuperTip = resources.GetString("RibEdgeRefreshAll.SuperTip");
            this.RibEdgeRefreshAll.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibEdgeRefreshAll_OnClick);
            // 
            // adxRibbonGroup3
            // 
            this.adxRibbonGroup3.Caption = "Run / Submit";
            this.adxRibbonGroup3.Controls.Add(this.RibEdgeParamRefresh);
            this.adxRibbonGroup3.Controls.Add(this.RibEdgeParamRefreshBook);
            this.adxRibbonGroup3.Id = "adxRibbonGroup_ef5eaf2e83894427a0276546ce85bc07";
            this.adxRibbonGroup3.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup3.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // RibEdgeParamRefresh
            // 
            this.RibEdgeParamRefresh.Caption = "Sheet";
            this.RibEdgeParamRefresh.Id = "adxRibbonButton_2ac952b8eeaa4f9f8509546703fa70c4";
            this.RibEdgeParamRefresh.Image = 4;
            this.RibEdgeParamRefresh.ImageList = this.ImageList_32X32;
            this.RibEdgeParamRefresh.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.RibEdgeParamRefresh.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibEdgeParamRefresh.ScreenTip = "Worksheet";
            this.RibEdgeParamRefresh.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.RibEdgeParamRefresh.SuperTip = "Displays reports parameter section in UI.";
            this.RibEdgeParamRefresh.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibEdgeParamRefresh_OnClick);
            // 
            // RibEdgeParamRefreshBook
            // 
            this.RibEdgeParamRefreshBook.Caption = "Book";
            this.RibEdgeParamRefreshBook.Id = "adxRibbonButton_147bc5eca31440d0a2dc449b44de0fe2";
            this.RibEdgeParamRefreshBook.Image = 5;
            this.RibEdgeParamRefreshBook.ImageList = this.ImageList_32X32;
            this.RibEdgeParamRefreshBook.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.RibEdgeParamRefreshBook.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibEdgeParamRefreshBook.ScreenTip = "Workbook";
            this.RibEdgeParamRefreshBook.SuperTip = "Displays reports parameter section of all reports except scheduled reports in UI." +
    "";
            this.RibEdgeParamRefreshBook.Visible = false;
            this.RibEdgeParamRefreshBook.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibEdgeParamRefreshBook_OnClick);
            // 
            // adxRibbonGroup4
            // 
            this.adxRibbonGroup4.Caption = "Run";
            this.adxRibbonGroup4.Controls.Add(this.RibEdgeShowHide);
            this.adxRibbonGroup4.Id = "adxRibbonGroup_d7f09161791e4be19cb16f397f717510";
            this.adxRibbonGroup4.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup4.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // RibEdgeShowHide
            // 
            this.RibEdgeShowHide.Caption = "Reports";
            this.RibEdgeShowHide.Id = "adxRibbonButton_198384d51bd246c6b2b209e225cf90b0";
            this.RibEdgeShowHide.Image = 6;
            this.RibEdgeShowHide.ImageList = this.ImageList_32X32;
            this.RibEdgeShowHide.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.RibEdgeShowHide.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibEdgeShowHide.ScreenTip = "Open UI";
            this.RibEdgeShowHide.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.RibEdgeShowHide.SuperTip = "Toggles the report\'s explorer window on/off.";
            this.RibEdgeShowHide.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibEdgeShowHide_OnClick);
            // 
            // adxRibbonGroup5
            // 
            this.adxRibbonGroup5.Caption = "Debug";
            this.adxRibbonGroup5.Controls.Add(this.RibEdgeDebug);
            this.adxRibbonGroup5.Controls.Add(this.RibEdgeIncludeOutputData);
            this.adxRibbonGroup5.Id = "adxRibbonGroup_8cfdf3b5c4884dac9197c521145c58a1";
            this.adxRibbonGroup5.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup5.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // RibEdgeDebug
            // 
            this.RibEdgeDebug.Caption = "Enable";
            this.RibEdgeDebug.Id = "adxRibbonCheckBox_171597373b2d4b239fbd50faff005df5";
            this.RibEdgeDebug.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibEdgeDebug.ScreenTip = "Debug Option";
            this.RibEdgeDebug.SuperTip = "Check the option to enable debugging.";
            this.RibEdgeDebug.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibEdgeDebug_OnClick);
            // 
            // RibEdgeIncludeOutputData
            // 
            this.RibEdgeIncludeOutputData.Caption = "Include Output Data";
            this.RibEdgeIncludeOutputData.Id = "adxRibbonCheckBox_dd81756b1a374d689a8de6b7672dc43d";
            this.RibEdgeIncludeOutputData.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibEdgeIncludeOutputData.ScreenTip = "Debug";
            this.RibEdgeIncludeOutputData.SuperTip = "When checked this will add the output data to the log file. Bydefault this is unc" +
    "hecked and no output data is printed in log file.";
            this.RibEdgeIncludeOutputData.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibEdgeIncludeOutputData_OnClick);
            // 
            // adxRibbonGroup6
            // 
            this.adxRibbonGroup6.Caption = "Settings";
            this.adxRibbonGroup6.Controls.Add(this.RibEdgeOptions);
            this.adxRibbonGroup6.Controls.Add(this.RibControlSheet);
            this.adxRibbonGroup6.Id = "adxRibbonGroup_eb7e34239b4b404cafa730cf64dd8432";
            this.adxRibbonGroup6.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup6.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // RibEdgeOptions
            // 
            this.RibEdgeOptions.Caption = "Options";
            this.RibEdgeOptions.Id = "adxRibbonButton_0372eca803a74d17b8c0d3860f16b708";
            this.RibEdgeOptions.Image = 7;
            this.RibEdgeOptions.ImageList = this.ImageList_32X32;
            this.RibEdgeOptions.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.RibEdgeOptions.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibEdgeOptions.ScreenTip = "User preferences";
            this.RibEdgeOptions.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.RibEdgeOptions.SuperTip = "Allows the users to select the preferences and save per session or per user.";
            this.RibEdgeOptions.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibEdgeOptions_OnClick);
            // 
            // RibControlSheet
            // 
            this.RibControlSheet.Caption = "Create Control Sheet";
            this.RibControlSheet.Id = "adxRibbonButton_ec3497f6e3de450cb25fb89f67c65fdf";
            this.RibControlSheet.Image = 8;
            this.RibControlSheet.ImageList = this.ImageList_32X32;
            this.RibControlSheet.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.RibControlSheet.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibControlSheet.ScreenTip = "Parameter Control Sheet";
            this.RibControlSheet.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.RibControlSheet.SuperTip = resources.GetString("RibControlSheet.SuperTip");
            this.RibControlSheet.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibControlSheet_OnClick);
            // 
            // adxRibbonGroup7
            // 
            this.adxRibbonGroup7.Caption = "About";
            this.adxRibbonGroup7.Controls.Add(this.RibEdgeAbout);
            this.adxRibbonGroup7.Id = "adxRibbonGroup_204f7b9b2b54441b8c56d4b9bcf5a961";
            this.adxRibbonGroup7.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup7.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // RibEdgeAbout
            // 
            this.RibEdgeAbout.Caption = "About";
            this.RibEdgeAbout.Id = "adxRibbonButton_21f28469752c4e48a1b47b607dd05dde";
            this.RibEdgeAbout.Image = 9;
            this.RibEdgeAbout.ImageList = this.ImageList_32X32;
            this.RibEdgeAbout.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.RibEdgeAbout.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibEdgeAbout.ScreenTip = "About";
            this.RibEdgeAbout.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.RibEdgeAbout.SuperTip = "About company and XLEdge product.";
            this.RibEdgeAbout.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibEdgeAbout_OnClick);
            // 
            // adxRibbonGroup8
            // 
            this.adxRibbonGroup8.Caption = "Help";
            this.adxRibbonGroup8.Controls.Add(this.RibEdgeHelp);
            this.adxRibbonGroup8.Id = "adxRibbonGroup_635077ae4340466bb36ed1c79c6a8f10";
            this.adxRibbonGroup8.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup8.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // RibEdgeHelp
            // 
            this.RibEdgeHelp.Caption = "Help";
            this.RibEdgeHelp.Id = "adxRibbonButton_b8c457afe865411383b10432868af110";
            this.RibEdgeHelp.Image = 10;
            this.RibEdgeHelp.ImageList = this.ImageList_32X32;
            this.RibEdgeHelp.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.RibEdgeHelp.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibEdgeHelp.ScreenTip = "Help";
            this.RibEdgeHelp.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.RibEdgeHelp.SuperTip = "Documentation for understanding XLEdge.";
            this.RibEdgeHelp.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibEdgeHelp_OnClick);
            // 
            // adxRibbonGroup9
            // 
            this.adxRibbonGroup9.Caption = " ";
            this.adxRibbonGroup9.Controls.Add(this.RibSheetLabel);
            this.adxRibbonGroup9.Id = "adxRibbonGroup_c9f93c3917324e8481f8be937836420a";
            this.adxRibbonGroup9.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup9.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // RibSheetLabel
            // 
            this.RibSheetLabel.Caption = " ";
            this.RibSheetLabel.Id = "adxRibbonLabel_6f3496d47e1c4d1583dcfd7129782dcc";
            this.RibSheetLabel.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxExcelTaskPanesManager1
            // 
            this.adxExcelTaskPanesManager1.Items.Add(this.adxExcelTaskPanesCollectionItem1);
            this.adxExcelTaskPanesManager1.SetOwner(this);
            // 
            // adxExcelTaskPanesCollectionItem1
            // 
            this.adxExcelTaskPanesCollectionItem1.IsHiddenStateAllowed = false;
            this.adxExcelTaskPanesCollectionItem1.IsMinimizedStateAllowed = false;
            this.adxExcelTaskPanesCollectionItem1.TaskPaneClassName = "XLEdge.ADXExcelTaskPane1";
            // 
            // adxExcelAppEvents1
            // 
            this.adxExcelAppEvents1.SheetSelectionChange += new AddinExpress.MSO.ADXExcelSheet_EventHandler(this.adxExcelAppEvents1_SheetSelectionChange);
            this.adxExcelAppEvents1.SheetBeforeDoubleClick += new AddinExpress.MSO.ADXExcelSheetBefore_EventHandler(this.adxExcelAppEvents1_SheetBeforeDoubleClick);
            this.adxExcelAppEvents1.SheetActivate += new AddinExpress.MSO.ADXHostActiveObject_EventHandler(this.adxExcelAppEvents1_SheetActivate);
            this.adxExcelAppEvents1.WorkbookActivate += new AddinExpress.MSO.ADXHostActiveObject_EventHandler(this.adxExcelAppEvents1_WorkbookActivate);
            this.adxExcelAppEvents1.SheetFollowHyperlink += new AddinExpress.MSO.ADXExcelHyperlink_EventHandler(this.adxExcelAppEvents1_SheetFollowHyperlink);
            this.adxExcelAppEvents1.SheetBeforeDelete += new AddinExpress.MSO.ADXExcelSheetBeforeDelete_EventHandler(this.AdxExcelAppEvents1_SheetBeforeDelete);
            // 
            // AddinModule
            // 
            this.AddinName = "XLEdge";
            this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaExcel;
            this.AddinStartupComplete += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinStartupComplete);
            this.AddinBeginShutdown += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinBeginShutdown);
            this.OnError += new AddinExpress.MSO.ADXError_EventHandler(this.AddinModule_OnError);
            this.OnRibbonLoaded += new AddinExpress.MSO.ADXRibbonLoaded_EventHandler(this.AddinModule_OnRibbonLoaded);

        }
        #endregion

        private AddinExpress.MSO.ADXRibbonTab TabXLEdge;
        internal System.Windows.Forms.ImageList ImageList_32X32;
        private AddinExpress.XL.ADXExcelTaskPanesManager adxExcelTaskPanesManager1;
        public AddinExpress.XL.ADXExcelTaskPanesCollectionItem adxExcelTaskPanesCollectionItem1;
        private AddinExpress.MSO.ADXExcelAppEvents adxExcelAppEvents1;
        private AddinExpress.MSO.ADXRibbonButton RibEdgeLogin;
        private AddinExpress.MSO.ADXRibbonButton RibEdgeLogout;
        private AddinExpress.MSO.ADXRibbonLabel RibLoginURL;
        private AddinExpress.MSO.ADXRibbonDialogBoxLauncher RibEdgeDialogBoxLauncher;
        private AddinExpress.MSO.ADXRibbonButton RibEdgeRefresh;
        private AddinExpress.MSO.ADXRibbonButton RibEdgeRefreshAll;
        private AddinExpress.MSO.ADXRibbonButton RibEdgeParamRefresh;
        private AddinExpress.MSO.ADXRibbonButton RibEdgeParamRefreshBook;
        private AddinExpress.MSO.ADXRibbonButton RibEdgeShowHide;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup2;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup3;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup4;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup5;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup6;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup7;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup8;
        private AddinExpress.MSO.ADXRibbonCheckBox RibEdgeDebug;
        private AddinExpress.MSO.ADXRibbonCheckBox RibEdgeIncludeOutputData;
        private AddinExpress.MSO.ADXRibbonButton RibEdgeOptions;
        private AddinExpress.MSO.ADXRibbonButton RibControlSheet;
        private AddinExpress.MSO.ADXRibbonButton RibEdgeAbout;
        private AddinExpress.MSO.ADXRibbonButton RibEdgeHelp;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup9;
        private AddinExpress.MSO.ADXRibbonLabel RibSheetLabel;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup1;
    }
}

