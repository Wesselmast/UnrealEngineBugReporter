using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Permissions;
using System.Threading.Tasks;

namespace SheetAPI {
    class Program {
        static readonly string[] scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string applicationName = "SheetTest";
        static readonly string spreadsheetID = "1mKnGo2BzPGxpC59Q7ZBAjuSeMHrIqOVJOWV4NTRPaY4";
        static readonly string sheet = "Entries";
        static SheetsService service;

        private static DateTime lastRead = DateTime.MinValue;
        private static string fileName = "bugsheet.txt";

        static void Main(string[] args) {
            GoogleCredential credential;

            Console.WriteLine("Welcome to the Unreal Engine 4 bug reporter! \n");
            Console.WriteLine("Initialising...\n");

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read)) {
                credential = GoogleCredential.FromStream(stream).CreateScoped(scopes);
            }

            service = new SheetsService(new BaseClientService.Initializer() {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });

            FileStream t = File.Create(fileName);
            t.Close();

            BindWatcher();
        }

        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
        static void BindWatcher() {
            using (FileSystemWatcher watcher = new FileSystemWatcher()) {
                watcher.Path = Directory.GetCurrentDirectory();
                watcher.NotifyFilter = NotifyFilters.LastWrite;

                watcher.Filter = fileName;
                watcher.Changed += OnChanged;

                watcher.EnableRaisingEvents = true;

                Console.WriteLine("Ready for input from Unreal Engine 4!");
                Console.WriteLine("Press 'q' to quit application");
                while (Console.Read() != 'q') ;
            }
        }

        static void OnChanged(object source, FileSystemEventArgs e) {
            DateTime lastWriteTime = File.GetLastWriteTime(fileName);
            if (lastWriteTime != lastRead) {
                int index = LastBugIndex() + 1;
                List<string> lines = new List<string>();
                lines.Add(index.ToString());
                lines.AddRange(File.ReadAllLines(fileName));
                AddBug(lines.ToArray());
                lastRead = lastWriteTime;
            }
        }

        static int LastBugIndex() {
            string range = $"{sheet}!A2:A";
            var request = service.Spreadsheets.Values.Get(spreadsheetID, range);

            ValueRange response = request.Execute();
            IList<IList<object>> values = response.Values;
            if (values == null || values.Count == 0) return 0;
            return int.Parse((string)response.Values[response.Values.Count - 1][0]);
        }

        static void AddBug(params string[] entry) {
            string range = $"{sheet}!A:F";
            ValueRange valueRange = new ValueRange();

            List<object> objectList = new List<object>();
            foreach (string row in entry) {
                objectList.Add(row);
            }

            valueRange.Values = new List<IList<object>> { objectList };
            var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }





        static void Fun(int amt) {
            Random rnd = new Random();
            string[] descriptions = {
                "Missing texture",
                "Not working",
                "Broken game",
                "Messy AI",
                "Dumb developers",
                "Some artifacts",
                "Bad game design",
                "Just not great",
                "Big lag spike here"
            };

            int lastBugIndex = LastBugIndex() + 1;

            for (int i = lastBugIndex; i < lastBugIndex + amt; i++) {
                string index = i.ToString();
                string prio = rnd.Next() % 2 == 0 ? "High priority" : "Low priority";
                string description = descriptions[rnd.Next() % descriptions.Length];
                string location = rnd.Next(1, 1000).ToString() + ", " +
                                  rnd.Next(1, 360).ToString() + ", " +
                                  rnd.Next(1, 1000).ToString();
                AddBug(i.ToString(), prio, description, location);
            }
        }
    }
}
