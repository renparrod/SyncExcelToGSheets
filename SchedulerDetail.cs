//using System;
//using SQLite.Net.Attributes;

namespace SyncExcelToGSheets
{
    //[Table("SchedulerDetail")]
    public class SchedulerDetail
    {
        //[PrimaryKey]
        public int SchedulerDetailId { get; set; }
        public int SchedulerId { get; set; }
        public int SourceDetailTypeId { get; set; }
        public int DestinationDetailTypeId { get; set; }
        public string SourceServer { get; set; }
        public string DestinationServer { get; set; }
        public string SourceData { get; set; }
        public string DestinationData { get; set; }
        public string SourceContainer { get; set; }
        public string DestinationContainer { get; set; }
    }
}
