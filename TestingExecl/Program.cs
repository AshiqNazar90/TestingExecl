using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestingExecl
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Create a list of accounts.
            var TestAccounts = new List<Account>
              {
            new Account {
                 Code ="001",
                 Name="John",
                 Description="Testing",
                 CreatedBy=01012022,
                 ModifiedBy=12012022,
                 EventName="Report",
                 Location="Uae",
                 Load=234.456
                },
             new Account {
                 Code ="002",
                 Name="Derick",
                 Description="Rule",
                 CreatedBy=2021,
                 ModifiedBy=2022,
                 EventName="Report",
                 Location="India",
                 Load=78.456
                }
             };

            // Display the list in an Excel spreadsheet.
            DisplayInExcel(TestAccounts);

            Console.WriteLine("Success");
            Console.ReadKey();
        }
        public class Account
        {


            public string Code { get; set; }
            public string Name { get; set; }
            public string Description { get; set; }
            public long CreatedBy { get; set; }
            public long ModifiedBy { get; set; }
   
            public string EventName { get; set; }
            public string Location { get; set; }
            public double Load { get; set; }
          

        }
        static void DisplayInExcel(IEnumerable<Account> accounts)
        {
            var excelApp = new Excel.Application();
       
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            workSheet.Cells[1, "A"] = "Code";
            workSheet.Cells[1, "B"] = "Name";
            workSheet.Cells[1, "C"] = "Description";
            workSheet.Cells[1, "D"] = "CreatedBy";
            workSheet.Cells[1, "E"] = "ModifiedBy";
            workSheet.Cells[1, "F"] = "EventName";
            workSheet.Cells[1, "G"] = "Location";
            workSheet.Cells[1, "H"] = "Load";
            var row = 1;
            foreach (var acct in accounts)
            {
                row++;
                workSheet.Cells[row, "A"] = acct.Code;
                workSheet.Cells[row, "B"] = acct.Name;
                workSheet.Cells[row, "C"] = acct.Description;
                workSheet.Cells[row, "D"] = acct.CreatedBy;
                workSheet.Cells[row, "E"] = acct.ModifiedBy;
                workSheet.Cells[row, "F"] = acct.EventName;
                workSheet.Cells[row, "G"] = acct.Location;
                workSheet.Cells[row, "H"] = acct.Load;

                workSheet.Columns[1].AutoFit();
                workSheet.Columns[2].AutoFit();
                workSheet.Range["A1", "H3"].AutoFormat(
                    Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                workSheet.Range["A1:B3"].Copy();

            }
        }


    }
}
