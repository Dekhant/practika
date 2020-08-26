using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;

namespace WindowsFormsApp2
{
    public class logic
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Sacha";
        static readonly string SpreadsheetId = "1RGGGdqUb-O137YLCGbrMvEBb7MslNqEiAWn4yE5R6bE";
        static readonly string sheet = "Class Data";
        static SheetsService service;
        public void initializeDates(DateTime start, DateTime end, ref List<string> dates)
        {
            string formatted = start.ToString("dd-MM-yyyy");
            DateTime mayDay;
            while (start.ToString("dd-MM-yyyy") != end.ToString("dd-MM-yyyy"))
            {
                formatted = start.ToString("dd-MM-yyyy");
                dates.Add(formatted);

                mayDay = start.AddDays(1);
                start = mayDay;
            }
            formatted = start.ToString("dd-MM-yyyy");
            dates.Add(formatted);
        }


        public int tableMarkup(string numOfCategories, string[] numOfCategory, string[] roomCapacity)
        {
            int numOfSeats = 0;
            var category = new List<int>();
            for (var i = 0; i < numOfCategory.Length; i++)
            {
                category.Add(int.Parse(numOfCategory[i]));
                numOfSeats += category[i] * int.Parse(roomCapacity[i]);
            }

            return numOfSeats;
        }

        public void updateRequest(string range, List<object> objectList)
        {
            GoogleCredential credential;
            using (var stream = new FileStream("My First Project-2dfed0050064.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            ValueRange valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { objectList };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
        }
    }
}
