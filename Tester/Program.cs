using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Tester
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Console.WriteLine(STC.Automation.Office.Excel.Utilities.Ranges.ConvertFormat("C"));

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
