using System;

namespace TrelloAutomation
{
    class TrelloCredentials
    {
        public static string ReadAPIKey()
        {
            Console.WriteLine("Enter your Trello API key:");
            return Console.ReadLine();
        }

        public static string ReadToken()
        {
            Console.WriteLine("Enter your Trello token:");
            return Console.ReadLine();
        }

        public static string ReadOrganizationId()
        {
            Console.WriteLine("Enter OrganizationId:");
            // OrganizationId or teamId can be obtained by typing .json at the end of the trello url
            return Console.ReadLine();
        }
    }
}
