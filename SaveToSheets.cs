using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SyncExcelToGSheets
{
    public partial class SaveToSheets
    {
        private string GoogleSheetId;
        private string GoogleSheetName;
        private string GoogleSheetRange;
        private string ExcelSheetName;
        private string ExcelSheetRange;
        private bool EnableSaveSheets;

        private void SaveToSheets_Load(object sender, RibbonUIEventArgs e)
        {
            GoogleSheetId = ConfigValues.GoogleSheetId;
            GoogleSheetName = ConfigValues.GoogleSheetName;
            GoogleSheetRange = ConfigValues.GoogleSheetRange;
            ExcelSheetName = ConfigValues.ExcelSheetName;
            ExcelSheetRange = ConfigValues.ExcelSheetRange;
            EnableSaveSheets = ConfigValues.EnableSaveSheets;
            EnableSync.Label = EnableSync.Checked ? "Deshabiltar Google Sheets" : "Habiltar Google Sheets";
        }

        private void EnableSync_Click(object sender, RibbonControlEventArgs e)
        {
            if (string.IsNullOrEmpty(GoogleSheetName) || string.IsNullOrEmpty(GoogleSheetId) || string.IsNullOrEmpty(ExcelSheetName))
            {
                var formConfig = new ConfigSync();
                formConfig.ShowDialog();
            }

            ConfigSyncSheet.Enabled = EnableSync.Checked;
            EnableSync.Label = EnableSync.Checked ? "Deshabiltar Google Sheets" : "Habiltar Google Sheets";

            ConfigValues.EnableSaveSheets = EnableSync.Checked;
        }

        private void ConfigSyncSheet_Click(object sender, RibbonControlEventArgs e)
        {
            var formConfig = new ConfigSync();
            //formConfig.GoogleSheetId { get; set; }
            //formConfig.GoogleSheetName { get; set; }
            //formConfig.GoogleSheetRange { get; set; }
            //formConfig.LocalSheetName { get; set; }
            //formConfig.LocalSheetRange { get; set; }
            //formConfig.EnableSaveSheets { get; set; }
            formConfig.ShowDialog();
        }
    }
}
