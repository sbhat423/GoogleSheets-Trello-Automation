using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Manatee.Trello;

namespace TrelloAutomation
{
    class TrelloOperator
    {
        public static async Task ListValuesToTrelloElements(IList<IList<Object>> values, Organization organization)
        {
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
                                    board = await CreateBoard(organization, cellValue);
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

        private static async Task<IBoard> CreateBoard(Organization org, string boardName)
        {
            var boardTask = await org.Boards.Add(boardName);
            return boardTask;
        }

        private static async Task<IList> AddList(IBoard board, string listName)
        {
            var listTask = await board.Lists.Add(listName);
            return listTask;
        }

        private static List<Member> CreateMemberList(string membersstring)
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

        private static async Task AddLabel(ICard card, string labelName, LabelColor colour)
        {
            await card.Board.Labels.Add(labelName, colour);
        }

        private static async Task AddChecklist(ICard card, string checklistName)
        {
            await card.CheckLists.Add(checklistName);
        }
    }
}
