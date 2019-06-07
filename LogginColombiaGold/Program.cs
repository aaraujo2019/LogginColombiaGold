using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Globalization;

namespace LogginColombiaGold
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
            Application.Run(new frmLogin());
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            //NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
            //nfi.CurrencyDecimalSeparator = ".";
            //nfi.CurrencyGroupSeparator = ",";
        }
    }
}
