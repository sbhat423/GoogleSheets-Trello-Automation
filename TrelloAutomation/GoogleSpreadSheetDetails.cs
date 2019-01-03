using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TrelloAutomation
{
    class GoogleSpreadSheetDetails
    {
        public static string ReadSpreadSheetId()
        {
            Console.WriteLine("Enter Spreadsheet ID:");
            return Console.ReadLine(); 
        }
        
        public static string ReadSpreadSheetRange()
        {
            Console.WriteLine("Enter the Range:");
            // Example "A1:H2"
            // TODO: Remove range input
            return Console.ReadLine(); ;

        }
    }
}
