using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.ComponentModel;

namespace SyncExcelToGSheets
{
    public partial class ThisAddIn
    {
        private BackgroundWorker SaveToSheetsWorker = new BackgroundWorker();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ConfigValues.ExcelApplication = Application;
            Application.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            //Application.WorkbookAfterSave += new Excel.AppEvents_WorkbookAfterSave(Application_WorkbookAfterSave);

            SaveToSheetsWorker.WorkerReportsProgress = true;
            SaveToSheetsWorker.DoWork += SaveToSheetsWorker_DoWork;
            //UpdateConfig();
        }

        private void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        //private void Application_WorkbookAfterSave(Excel.Workbook Wb, bool Success)
        {

            //TestProperties();
            //if (EnableSaveSheets)
            if (ConfigValues.EnableSaveSheets)
            {
                //Thread threadSave = new Thread(new ThreadStart(SaveDestinationData));
                //threadSave.Start();
                //WriteStatusBar("Leyendo datos desde Excel...");

                var taskParams = new DataParameters()
                {
                    DataForSync = GetDataFromExcel(),
                    GoogleSheetName = ConfigValues.GoogleSheetName,
                    GoogleSheetRange = ConfigValues.GoogleSheetRange,
                    SpreadsheetId = ConfigValues.GoogleSheetId
                };
                SaveToSheetsWorker.RunWorkerAsync(argument: taskParams);
            }
        }

        private List<IList<object>> GetDataFromExcel()
        {
            //WriteStatusBar("Escribiendo hacia Google Sheets");
            var dataForSheets = new List<IList<object>>();
            var excel = Application;
            var xlbook = excel.ActiveWorkbook;
            var worksheets = xlbook.Worksheets;

            try
            {
                // TODO: Verify if sheet exists
                var sheet = (Excel.Worksheet)worksheets[ConfigValues.ExcelSheetName];
                int row = 1;

                while (!string.IsNullOrEmpty(((Excel.Range)sheet.Cells[row, 1]).Value))
                {
                    var rowData = new List<object>();
                    int col = 1;

                    while (!string.IsNullOrEmpty(((Excel.Range)sheet.Cells[1, col]).Value))
                    {
                        rowData.Add(((Excel.Range)sheet.Cells[row, col]).Value);
                        col++;
                    }

                    dataForSheets.Add(rowData);
                    row++;
                }
            }
            catch { }

            return dataForSheets;
        }

        private void SaveToSheetsWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            //UpdateConfig();
            var taskParams = (DataParameters)e.Argument;
            //var dataForSheets = (List<IList<object>>)e.Argument;
            var dataForSheets = taskParams.DataForSync;
            //var dataForSheets = new List<IList<object>>();
            //var excel = Application;
            //var xlbook = excel.ActiveWorkbook;
            ////var worksheets = xlbook.Worksheets;

            try
            {
                //// TODO: Verify if sheet exists
                //var sheet = (Excel.Worksheet)worksheets[LocalSheetName];
                //int row = 1;

                //while (!string.IsNullOrEmpty(((Excel.Range)sheet.Cells[row, 1]).Value))
                //{
                //    var rowData = new List<object>();
                //    int col = 1;

                //    while (!string.IsNullOrEmpty(((Excel.Range)sheet.Cells[1, col]).Value))
                //    {
                //        rowData.Add(((Excel.Range)sheet.Cells[row, col]).Value);

                //        col++;
                //    }

                //    dataForSheets.Add(rowData);
                //    row++;
                //}


                var detail = new Dictionary<string, string>()
                    {
                        { "SpreadsheetId", taskParams.SpreadsheetId },
                        { "SheetName", taskParams.GoogleSheetName },
                        { "Range", taskParams.GoogleSheetRange }
                    };

                ExecuteTask("SyncExcelToGSheets", "GoogleSheetsService", detail, dataForSheets);
            }
            catch (Exception ex)
            {
                var d = ex.Message;
            }
        }

        private static void ExecuteTask(string assemblyName, string sourceType, Dictionary<string, string> detailProperties, List<IList<object>> dataToSave)
        {
            var sourceHandle = Activator.CreateInstance(assemblyName, $"{assemblyName}.{sourceType}");
            var sourceService = (IDataService)sourceHandle.Unwrap();
            sourceService.SourceDetail = detailProperties;
            sourceService.SaveDestinationData(dataToSave);
        }

        private string ReadDocumentProperty(string propertyName)
        {
            var excel = Application;
            var xlbook = excel.ActiveWorkbook;

            Office.DocumentProperties properties;
            properties = (Office.DocumentProperties)xlbook.CustomDocumentProperties;

            foreach (Office.DocumentProperty prop in properties)
            {
                if (prop.Name == propertyName)
                {
                    return prop.Value.ToString();
                }
            }
            return null;
        }

        private void WriteDocumentProperty(string propertyName, string propertyValue)
        {
            var excel = Application;
            var xlbook = excel.ActiveWorkbook;

            Office.DocumentProperties properties;
            properties = (Office.DocumentProperties)xlbook.CustomDocumentProperties;

            if (ReadDocumentProperty(propertyName) != null)
            {
                properties[propertyName].Delete();
            }

            properties.Add(propertyName, false,
                Office.MsoDocProperties.msoPropertyTypeString,
                propertyValue);
        }

        public List<string> GetAllSheets()
        {
            var sheets = new List<string>();

            foreach (Excel.Worksheet wsh in Application.Worksheets)
            {
                sheets.Add(wsh.Name);
            }

            return sheets;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
