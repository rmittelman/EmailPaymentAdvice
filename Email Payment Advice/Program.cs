using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Configuration;
using System.Data.Odbc;
using SendGrid;
using SendGrid.Helpers.Mail;
using Aimm.Logging;

namespace EmailPaymentAdvice
{
    static class Program
    {


        [STAThread]
        static void Main(string[] args)
        {
            LogIt.LogInfo("Program starting");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        private static void SetPermissions()
        {
            throw new NotImplementedException();
        }
    }
}
