using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Reliance
{
    static class Program
    {
     [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            var statusFrom = new Status();
            Application.Run(statusFrom);

            //RBot bot = new RBot("chrome");
            //bot.RunTheBot("kartik.rpa@gmail.com", "Kartik1a$");
        }
    }
}
