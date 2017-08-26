using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Swap.model
{
    class Record
    {
        public string caseNo { get; set; }
        public string billDate { get; set; }
        public string billEventCode { get; set; }
        public string billRates { get; set; }
        public string billUnits { get; set; }
        public string clientID { get; set; }
        public string location { get; set; }
        public string comment { get; set; }
        public Record(string a, string b, string c, string d, string e, string f, string g, string h)
        {
            caseNo = a;
            billDate = b;
            billEventCode = c;
            billRates = d;
            billUnits = e;
            clientID = f;
            location = g;
            comment = h;
        }
        public void printData()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine(caseNo + "\t" + billDate + "\t" + billEventCode + "\t" + billRates + "\t" + billUnits + "\t" + clientID + "\t" + location + "\t" + comment);
            Console.ForegroundColor = ConsoleColor.White;
        }
    }
}
