using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLSample
{
    class Program
    {
        static void Main(string[] args)
        {
            Report report = new Report();

            report.CreateExcelDoc(@"D:\Report.xlsx");

            Console.WriteLine("Excel file has created!");
        }
    }
}
