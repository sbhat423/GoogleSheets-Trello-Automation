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

namespace CSV_to_Trello
{
    class TrelloCardsGenerator
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";
        private String organizationId;
        private String googleSpreadSheetId;
        private String googleSpreadSheetRange;
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

        private async Task<IBoard> CreateBoard(Organization org, String boardName)
        {
            var boardTask = await org.Boards.Add(boardName);
            return boardTask;
        }

        private async Task<IList> AddList(IBoard board, String listName)
        {
            var listTask = await board.Lists.Add(listName);
            return listTask;
        }

        private List<Member> CreateMemberList(String membersString)
        {
            var memberList = new List<Member>();
            String[] members = membersString.Split(';');
            foreach (String memberString in members)
            {
                var member = new Member(membersString);
                memberList.Add(member);
            }
            return memberList;
        }

        private async Task AddLabel(ICard card, String labelName, LabelColor colour)
        {
            await card.Board.Labels.Add(labelName, colour);
        }

        private async Task AddChecklist(ICard card, String checklistName)
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
            googleSpreadSheetRange = Console.ReadLine(); ;
        }

        private async Task ReadGoogleSpreadSheet()
        {
            UserCredential credential;
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
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
                    String title = null;
                    String description = null;
                    List<Member> memberList = null;
                    ICard card = null;
                    //DateTime dueDate;
                    int count = 0;
                    foreach (var cell in row)
                    {
                        if (readHeader == false)
                        {
                            keyList.Add(cell.ToString());
                        }
                        else
                        {
                            switch (keyList[count])
                            {
                                case "Board":
                                    Console.WriteLine("Creating new Board..........................");
                                    board = await CreateBoard(org, cell.ToString());
                                    break;
                                case "List":
                                    Console.WriteLine("Adding new List..........................");
                                    if (board != null)
                                    {
                                        await board.Lists.Refresh();
                                        list = await AddList(board, cell.ToString());
                                    }
                                    break;
                                case "Title":
                                    Console.WriteLine("Adding Titile to the card..........................");
                                    title = cell.ToString();
                                    card = await list.Cards.Add(title);
                                    break;
                                case "Description":
                                    Console.WriteLine("Adding description to the card..........................");
                                    description = cell.ToString();
                                    card.Description = description;
                                    break;
                                case "Member":
                                    Console.WriteLine("Adding member list to the card..........................");
                                    memberList = CreateMemberList(cell.ToString());
                                    foreach (Member member in memberList)
                                    {
                                        await card.Members.Add(member);
                                    }
                                    break;
                                case "DueDate":
                                    Console.WriteLine("Creating new Card..........................");
                                    card.DueDate = DateTime.Parse(cell.ToString());
                                    break;
                                case "Labels":
                                    await AddLabel(card, "Urgent", LabelColor.Red);
                                    break;
                                case "Checklist":
                                    await AddChecklist(card, "spraying");
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