using System;
using Google.Apis.Sheets.v4;
using System.Text.Json;
using System.Text.RegularExpressions;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4.Data;

namespace GSAL
{
    public class Program
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "GSAL";
        //https://docs.google.com/spreadsheets/d/1OIHhOo90A6L-7IOwujd0j71i-znGD6SZLaX4tiZUirE/edit?usp=sharing
        //
        static readonly string SpreadSheetId = "1OIHhOo90A6L-7IOwujd0j71i-znGD6SZLaX4tiZUirE";
        static readonly string sheet = "list1";
        static SheetsService service;

        //var id = GetSheetId("https://docs.google.com/spreadsheets/d/1OIHhOo90A6L-7IOwujd0j71i-znGD6SZLaX4tiZUirE/edit?usp=sharing");
        public static void Main(string[] args)
        {
            GoogleCredential credential;

            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });

            DeleteEntry();
        }

        static void ReadEntries()
        {
            var range = $"{sheet}!A1:F7";
            var request = service.Spreadsheets.Values.Get(SpreadSheetId, range);

            var response = request.Execute();
            var values = response.Values;

            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    foreach (var item in row)
                    {
                        Console.Write(item + "|");
                    }
                    Console.WriteLine();
                }
            }
            else
            {
                Console.WriteLine("No data");
            }
        }

        //Next methods need a permission
        static void CreateEntry()
        {
            var range = $"{sheet}!A:F";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { "Hello!", "This", "was", "inserted", "via", "c#" };
            valueRange.Values = new List<IList<object>> { objectList };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadSheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponce = appendRequest.Execute();
        }

        static void UpdateEntry()
        {
            var range = $"{sheet}!D8";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { "updated!"};
            valueRange.Values = new List<IList<object>> { objectList };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadSheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponce = updateRequest.Execute();

        }

        static void DeleteEntry()
        {
            var range = $"{sheet}!D8";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody,SpreadSheetId,range);
            var deleteResponce = deleteRequest.Execute();
        }
    }
}