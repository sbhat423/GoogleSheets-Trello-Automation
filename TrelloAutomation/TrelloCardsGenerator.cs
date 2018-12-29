using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Manatee.Trello;
using System.Threading.Tasks;

namespace TrelloAutomation
{
    class TrelloCardsGenerator
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart"; // TODO add to app.config
        private string organizationId;
        private string googleSpreadSheetId;
        private string googleSpreadSheetRange;
        private Organization org;

        static void Main(string[] args)
        {
            TrelloCardsGenerator trelloCardsGenerator = new TrelloCardsGenerator();
            trelloCardsGenerator.FetchTrelloAPICredentials();
            trelloCardsGenerator.FetchGoogleSpreadSheetDetails();
            trelloCardsGenerator.ReadGoogleSpreadSheet().Wait();
            Console.WriteLine("Press any key to exit");
            Console.ReadLine();

        }

        private async Task<IBoard> CreateBoard(Organization org, string boardName)
        {
            var boardTask = await org.Boards.Add(boardName);
            return boardTask;
        }

        private async Task<IList> AddList(IBoard board, string listName)
        {
            var listTask = await board.Lists.Add(listName);
            return listTask;
        }

        private List<Member> CreateMemberList(string membersstring)
        {
            var memberList = new List<Member>();
            string[] members = membersstring.Split(';');
            foreach (string memberstring in members)
            {
                var member = new Member(membersstring);
                memberList.Add(member);
            }
            return memberList;
        }

        private async Task AddLabel(ICard card, string labelName, LabelColor colour)
        {
            await card.Board.Labels.Add(labelName, colour);
        }

        private async Task AddChecklist(ICard card, string checklistName)
        {
            await card.CheckLists.Add(checklistName);
        }

        private void FetchTrelloAPICredentials()
        {
            Console.WriteLine("Enter your Trello API key:");
            TrelloAuthorization.Default.AppKey = Console.ReadLine();
            Console.WriteLine("Enter your Trello token:");
            TrelloAuthorization.Default.UserToken = Console.ReadLine(); ;
            Console.WriteLine("Enter OrganizationId:");
            // OrganizationId or teamId can be obtained by typing .json at the end of the trello url
            organizationId = Console.ReadLine();
            org = new Organization(organizationId);
        }

        private void FetchGoogleSpreadSheetDetails()
        {
            // Define request parameters.
            Console.WriteLine("Enter Spreadsheet ID:");
            googleSpreadSheetId = Console.ReadLine(); ;
            Console.WriteLine("Enter the Range:");
            // Example "A1:H2"
            // TODO: Remove range input
            googleSpreadSheetRange = Console.ReadLine(); ;
        }

        private async Task ReadGoogleSpreadSheet()
        {
            UserCredential credential;
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

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(googleSpreadSheetId, googleSpreadSheetRange);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            List<string> keyList = new List<string>();
            if (values != null && values.Count > 0)
            {
                bool readHeader = false;
                foreach (var row in values)
                {
                    IBoard board = null;
                    IList list = null;
                    string title = null;
                    string description = null;
                    List<Member> memberList = null;
                    ICard card = null;
                    int count = 0;
                    foreach (var cell in row)
                    {
                        string cellValue = cell.ToString();
                        if (readHeader == false)
                        {
                            keyList.Add(cellValue);
                        }
                        else
                        {
                            switch (keyList[count])
                            {
                                case "Board":
                                    Console.WriteLine("Creating new Board..........................");
                                    board = await CreateBoard(org, cellValue);
                                    break;
                                case "List":
                                    Console.WriteLine("Adding new List..........................");
                                    if (board != null)
                                    {
                                        await board.Lists.Refresh();
                                        list = await AddList(board, cellValue);
                                    }
                                    break;
                                case "Title":
                                    Console.WriteLine("Adding Titile to the card..........................");
                                    title = cellValue;
                                    card = await list.Cards.Add(title);
                                    break;
                                case "Description":
                                    Console.WriteLine("Adding description to the card..........................");
                                    description = cellValue;
                                    card.Description = description;
                                    break;
                                case "Member":
                                    Console.WriteLine("Adding member list to the card..........................");
                                    memberList = CreateMemberList(cellValue);
                                    foreach (Member member in memberList)
                                    {
                                        await card.Members.Add(member);
                                    }
                                    break;
                                case "DueDate":
                                    Console.WriteLine("Creating new Card..........................");
                                    card.DueDate = DateTime.Parse(cellValue);
                                    break;
                                case "Labels":
                                    await AddLabel(card, cellValue, LabelColor.Red);
                                    break;
                                case "Checklist":
                                    await AddChecklist(card, cellValue);
                                    break;
                                default:
                                    break;
                            }
                        }
                        count++;
                    }
                    count = 0;
                    readHeader = true;
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }
    }
}