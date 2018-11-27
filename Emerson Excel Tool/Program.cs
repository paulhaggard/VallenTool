using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Emerson_Excel_Tool
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new ToolForm());

            /*////
            /// Feature Waitlist
            /// Table data is trash.  Need to either manage with datatables or excel tables. try datatables first.
            /// Need to add statistical stuff like Ryan demonstrated.
            /// then, try charting data.
            /// Time left: organize, make prettier interface/more responsive.  
            /// Then get critiques!
            /// 
            /// Bug List
            /// 
            /////once, it closed another instance. not repeatable???
            */
        }
    }
}
