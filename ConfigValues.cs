using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.Office.Core;

namespace SyncExcelToGSheets
{
    public static class ConfigValues
    {
        public static string GoogleSheetId { get => ReadDocumentProperty("GoogleSheetId"); set { WriteDocumentProperty("GoogleSheetId", value); } }
        public static string GoogleSheetName { get => ReadDocumentProperty("GoogleSheetName"); set { WriteDocumentProperty("GoogleSheetName", value); } }
        public static string GoogleSheetRange { get => ReadDocumentProperty("GoogleSheetRange"); set { WriteDocumentProperty("GoogleSheetRange", value); } }
        public static string ExcelSheetName { get => ReadDocumentProperty("LocalSheetName"); set { WriteDocumentProperty("LocalSheetName", value); } }
        public static string ExcelSheetRange { get => ReadDocumentProperty("LocalSheetRange"); set { WriteDocumentProperty("LocalSheetRange", value); } }
        public static bool EnableSaveSheets { get => Convert.ToBoolean(ReadDocumentProperty("EnableSaveSheets")); set { WriteDocumentProperty("EnableSaveSheets", value.ToString()); } }


        public static Excel.Application ExcelApplication { get; set; }
        public static Excel.Workbook ActiveWorkbook { get; set; }



        //var excel = Application;
        //var xlbook = excel.ActiveWorkbook;
        //var worksheets = xlbook.Worksheets;

        private static string ReadDocumentProperty(string propertyName)
        {
            try
            {
                if (ExcelApplication != null && ExcelApplication.ActiveWorkbook != null)
                {
                    var xlbook = ExcelApplication.ActiveWorkbook;

                    //Office.DocumentProperties properties;
                    DocumentProperties properties = (DocumentProperties)xlbook.CustomDocumentProperties;

                    foreach (DocumentProperty prop in properties)
                    {
                        if (prop.Name == propertyName)
                        {
                            return prop.Value.ToString();
                        }
                    }
                }
            }
            catch
            {
                return null;
            }
            return null;
        }

        private static void WriteDocumentProperty(string propertyName, string propertyValue)
        {
            try
            {
                var xlbook = ExcelApplication.ActiveWorkbook;

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
            catch { }
        }
    }
}
