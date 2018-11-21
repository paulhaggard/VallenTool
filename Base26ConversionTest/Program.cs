using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToolkit;

namespace Base26ConversionTest
{
    class Program
    {
        static void Main(string[] args)
        {
            while(true)
            {
                Console.WriteLine("Please enter a number: ");
                string s;
                int num;
                bool first = true;

                do
                {
                    if (!first)
                        Console.WriteLine("Incorrect format, please try again: ");

                    s = Console.ReadLine();

                    if (first)
                        first = false;

                } while (!int.TryParse(s, out num));    // Repeat until s is parsable as an int

                Console.WriteLine("Base 26 value of {0} is {1}.", num, ExcelPort.ColumnNumToColumnString(num));
            }
        }
    }
}
