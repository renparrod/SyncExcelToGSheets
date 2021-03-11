using System.Collections.Generic;

namespace SyncExcelToGSheets
{
    public class DataParameters
    {
        public List<IList<object>> DataForSync { get; set; }
        public string SpreadsheetId { get; set; }
        public string GoogleSheetName { get; set; }
        public string GoogleSheetRange { get; set; }
    }
}
