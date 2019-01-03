using System.IO;
using Google.Apis.Auth.OAuth2;
using System.Threading;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;


namespace TrelloAutomation
{
    class GoogleSpreadSheetReader
    {
        private static UserCredential credential;
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart"; // TODO add to app.config

        public static void GetCredentials()
        {
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string googleCredentialsFilePath = "token.json"; // TODO add to app.config
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                                Scopes,
                                "user",
                                CancellationToken.None,
                                new FileDataStore(googleCredentialsFilePath, true)).Result;
            }
        }

        public static ValueRange ReadSheet(string googleSpreadSheetId, string googleSpreadSheetRange)
        {
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(googleSpreadSheetId, googleSpreadSheetRange);
            return request.Execute();
        }
    }
}
