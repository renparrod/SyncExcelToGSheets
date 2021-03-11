using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace SyncExcelToGSheets
{
    public class GoogleSheetsService : IDataService
    {
        public SchedulerDetail DetailScheduler { get; set; }
        //public LogService LoggingService { get; set; }
        public Dictionary<string, string> SourceDetail { get; set; }
        public Dictionary<string, string> DestinatioDetail { get; set; }

        string[] Scopes = { SheetsService.Scope.Spreadsheets };
        string ApplicationName = "Sync Excel to Google Sheets";
        ServiceAccountCredential _credential = null;
        SheetsService _sheetsService = null;

        // https://console.developers.google.com
        private string serviceAccountEmail = ""; // Change the username (User application in Google Cloud Platform)
        string credentialsFileName = "credentials.json"; // This file has to be generated in GCP, enable Google Sheet API



        public GoogleSheetsService()
        {
            string jsonFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, credentialsFileName);

            using (Stream stream = new FileStream(jsonFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                _credential = (ServiceAccountCredential)GoogleCredential.FromStream(stream).UnderlyingCredential;

                var initializer = new ServiceAccountCredential.Initializer(_credential.Id)
                {
                    User = serviceAccountEmail,
                    Key = _credential.Key,
                    Scopes = Scopes
                };
                _credential = new ServiceAccountCredential(initializer);
            }
        }

        private void SheetsServiceInitializer()
        {
            if (_sheetsService == null)
            {
                _sheetsService = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = _credential,
                    ApplicationName = ApplicationName,
                });
            }
        }

        public List<object> GetSourceData()
        {
            try
            {
                SheetsServiceInitializer();
                var _spreadsheetId = SourceDetail["SpreadsheetId"];
                var _range = $"{SourceDetail["SheetName"]}!{SourceDetail["Range"]}";

                var request = _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, _range);
                var response = request.Execute();
                // TODO: Convert to List<object>
                // return response.Values;
                return new List<object>();
            }
            catch //(Exception ex)
            {
                return new List<object>();
            }
        }

        //public async Task<int> SaveDestinationData(List<object> sourceData)
        //public void SaveDestinationData(List<object> sourceData)
        public void SaveDestinationData(List<IList<object>> sourceData)
        {
            try
            {
                SheetsServiceInitializer();
                var _spreadsheetId = SourceDetail["SpreadsheetId"];
                var _range = $"{SourceDetail["SheetName"]}!{SourceDetail["Range"]}";

                // var data = new List<IList<object>>();
                if (sourceData.Count > 0)
                {
                    //var list = (from i in (ExpandoObject)sourceData[0] select i.Key).Cast<object>().ToList();
                    //data.Add(list);

                    ClearAllValuesFromSheet(_spreadsheetId, 0);

                    //data.AddRange(sourceData.Select(item => (from kvp in (ExpandoObject)item select kvp.Value).ToList())
                    //                            .Cast<IList<object>>().ToList());

                    var valueInputOption = "USER_ENTERED";
                    var updateData = new List<ValueRange>();
                    var dataValueRange = new ValueRange
                    {
                        Range = _range,
                        //Values = data
                        Values = sourceData
                    };
                    updateData.Add(dataValueRange);

                    var requestBody = new BatchUpdateValuesRequest
                    {
                        ValueInputOption = valueInputOption,
                        Data = updateData
                    };

                    var requestBodyEmpty = new BatchUpdateValuesRequest
                    {
                        ValueInputOption = valueInputOption,
                        Data = null
                    };

                    var request = _sheetsService.Spreadsheets.Values.BatchUpdate(requestBodyEmpty, _spreadsheetId);
                    var response = request.Execute();

                    request = _sheetsService.Spreadsheets.Values.BatchUpdate(requestBody, _spreadsheetId);
                    response = request.Execute();
                }
            }
            catch //(Exception ex)
            {
            }
        }

        private void ClearAllValuesFromSheet(string spreadsheetId, int sheetId)
        {
            Request RequestBody = new Request()
            {
                UpdateCells = new UpdateCellsRequest()
                {
                    Range = new GridRange()
                    {
                        SheetId = sheetId
                    },
                    Fields = "*"
                }
            };

            var RequestContainer = new List<Request>();
            RequestContainer.Add(RequestBody);

            var clearRequest = new BatchUpdateSpreadsheetRequest();
            clearRequest.Requests = RequestContainer;

            var batchUpdateRequest = new SpreadsheetsResource.BatchUpdateRequest(_sheetsService, clearRequest, spreadsheetId);
            batchUpdateRequest.Execute();
        }

        public void AddNewRow(string spreadsheetId)
        {
            Request RequestBody = new Request()
            {
                InsertDimension = new InsertDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        SheetId = 0,
                        Dimension = "ROWS",
                        StartIndex = 0,
                        EndIndex = 10
                    }
                }
            };

            List<Request> RequestContainer = new List<Request>();
            RequestContainer.Add(RequestBody);

            BatchUpdateSpreadsheetRequest insertRequest = new BatchUpdateSpreadsheetRequest();
            insertRequest.Requests = RequestContainer;

            var insert = new SpreadsheetsResource.BatchUpdateRequest(_sheetsService, insertRequest, spreadsheetId);
            insert.Execute();
        }

        private string GetColumnName(int columnIndex)
        {
            var div = columnIndex;
            var col = string.Empty;
            int mod;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                col = Convert.ToChar(65 + mod).ToString() + col;
                div = (div - mod) / 26;
            }

            return col;
        }

        public List<IList<object>> GetDataForGoogleSheets(List<object> dataReportSource)
        {
            var returnData = new List<IList<object>>();
            if (dataReportSource?.Count > 0)
            {
                returnData.AddRange(dataReportSource.Select(item => item.GetType().GetProperties()).Select(props => props.Select(pInfo => pInfo.Name).Cast<object>().ToList()));

                foreach (var item in dataReportSource)
                {
                    var props = item.GetType().GetProperties();
                    returnData.Add(props.Select(pInfo => pInfo.Name).Cast<object>().ToList());
                }
            }

            return returnData;
        }
    }
}