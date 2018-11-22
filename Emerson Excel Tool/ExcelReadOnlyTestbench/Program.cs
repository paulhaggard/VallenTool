using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReadOnlyTestbench
{
    class Program
    {
        static void Main(string[] args)
        {
            print("Opening Excel...");
            Excel.Application app = null;
            Excel._Workbook oWB;
            string filenameS = "vallen.xlsx";

            try
            {
                try
                {
                    // Check if Excel is already open
                    app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    print("Found an open version of excel.");
                }
                catch
                {
                    // If not create a new instance
                    app = new Excel.Application();
                    print("Created a new instance of excel.");
                }

                app.Visible = true;
                app.DisplayAlerts = false;  // Suppresses the save prompt after closing excel.

                try
                {
                    oWB = (app.Workbooks.Open(filenameS));
                    print("Opened the workbook from a previous file.");
                }
                catch (COMException)
                {
                    print("Creating a new workbook.");
                    oWB = app.Workbooks.Add();
                    oWB.SaveAs(filenameS);
                }

                Console.ReadKey();

                app.Quit();
            }
            catch (Exception)
            {
                // Makes sure to reset the variables in case they might have gotten set before the exception occured
                print("Unable to open excel, exitting...");
    
            }
        }

        static void print(string s)
        {
            Console.WriteLine(s);
        }
    }
}
