using System.Collections.Generic;

namespace SyncExcelToGSheets
{
    public interface IDataService
    {
        SchedulerDetail DetailScheduler { get; set; }
        Dictionary<string, string> SourceDetail { get; set; }
        Dictionary<string, string> DestinatioDetail { get; set; }
        List<object> GetSourceData();
        void SaveDestinationData(List<IList<object>> data);
    }
}
