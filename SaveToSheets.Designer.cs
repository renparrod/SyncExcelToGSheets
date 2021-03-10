
namespace SyncExcelToGSheets
{
    partial class SaveToSheets : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SaveToSheets()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.TabGoogleSheets = this.Factory.CreateRibbonTab();
            this.GroupConfigSheet = this.Factory.CreateRibbonGroup();
            this.EnableSync = this.Factory.CreateRibbonToggleButton();
            this.ConfigSyncSheet = this.Factory.CreateRibbonButton();
            this.TabGoogleSheets.SuspendLayout();
            this.GroupConfigSheet.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabGoogleSheets
            // 
            this.TabGoogleSheets.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabGoogleSheets.Groups.Add(this.GroupConfigSheet);
            this.TabGoogleSheets.Label = "Google Sheets";
            this.TabGoogleSheets.Name = "TabGoogleSheets";
            // 
            // GroupConfigSheet
            // 
            this.GroupConfigSheet.Items.Add(this.EnableSync);
            this.GroupConfigSheet.Items.Add(this.ConfigSyncSheet);
            this.GroupConfigSheet.Label = "Configuración Google Sheets";
            this.GroupConfigSheet.Name = "GroupConfigSheet";
            // 
            // EnableSync
            // 
            this.EnableSync.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.EnableSync.Image = global::SyncExcelToGSheets.Properties.Resources.GoogleSheetsLogo;
            this.EnableSync.Label = "Habilitar sincronización";
            this.EnableSync.Name = "EnableSync";
            this.EnableSync.ShowImage = true;
            this.EnableSync.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EnableSync_Click);
            // 
            // ConfigSyncSheet
            // 
            this.ConfigSyncSheet.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ConfigSyncSheet.Image = global::SyncExcelToGSheets.Properties.Resources.Config;
            this.ConfigSyncSheet.Label = "Configurar sincronización";
            this.ConfigSyncSheet.Name = "ConfigSyncSheet";
            this.ConfigSyncSheet.ShowImage = true;
            this.ConfigSyncSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ConfigSyncSheet_Click);
            // 
            // SaveToSheets
            // 
            this.Name = "SaveToSheets";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.TabGoogleSheets);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SaveToSheets_Load);
            this.TabGoogleSheets.ResumeLayout(false);
            this.TabGoogleSheets.PerformLayout();
            this.GroupConfigSheet.ResumeLayout(false);
            this.GroupConfigSheet.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabGoogleSheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GroupConfigSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton EnableSync;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ConfigSyncSheet;
    }

    partial class ThisRibbonCollection
    {
        internal SaveToSheets SaveToSheets
        {
            get { return this.GetRibbon<SaveToSheets>(); }
        }
    }
}
