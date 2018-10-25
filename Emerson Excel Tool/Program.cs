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

            /*/Bug List
            /// If Vallen file is closed before 'process' is run, app crashes.
            ///Similar - if Vallen file is closed before app is closed, app crashes.
            */
        }
    }
}
