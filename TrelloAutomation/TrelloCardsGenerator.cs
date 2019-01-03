using Google.Apis.Sheets.v4.Data;
using System;
using Manatee.Trello;
using System.Threading.Tasks;

namespace TrelloAutomation
{
    class TrelloCardsGenerator
    {
        private string organizationId;
        private string googleSpreadSheetId;
        private string googleSpreadSheetRange;
        private Organization trelloOrganization;
        private ValueRange response;

        static void Main(string[] args)
        {
            TrelloCardsGenerator trelloCardsGenerator = new TrelloCardsGenerator();
            trelloCardsGenerator.FetchTrelloAPICredentials();
            trelloCardsGenerator.FetchGoogleSpreadSheetDetails();
            trelloCardsGenerator.ReadGoogleSpreadSheet();
            trelloCardsGenerator.GenerateTrelloElements().Wait();
            Console.Write("Press any key to exit...");
            Console.ReadLine();
        }

        private void FetchTrelloAPICredentials()
        {
            TrelloAuthorization.Default.AppKey = TrelloCredentials.ReadAPIKey();
            TrelloAuthorization.Default.UserToken = TrelloCredentials.ReadToken();
            organizationId = TrelloCredentials.ReadOrganizationId();
            trelloOrganization = new Organization(organizationId);
        }

        private void FetchGoogleSpreadSheetDetails()
        {
            googleSpreadSheetId = GoogleSpreadSheetDetails.ReadSpreadSheetId();
            googleSpreadSheetRange = GoogleSpreadSheetDetails.ReadSpreadSheetRange();
        }

        private void ReadGoogleSpreadSheet()
        {
            GoogleSpreadSheetReader.GetCredentials();
            response = GoogleSpreadSheetReader.ReadSheet(googleSpreadSheetId, googleSpreadSheetRange);
        }

        private async Task GenerateTrelloElements()
        {
            await TrelloOperator.ListValuesToTrelloElements(response.Values, trelloOrganization);
        }
    }
}