using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToolkit;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPort port = new ExcelPort();
            port.setVisible(true);
            port.CreateWorkbook("vallen.xlsx");
            Console.WriteLine("Excel should now be running");
            Console.Read();
            port.CloseApp();
        }
    }
}
