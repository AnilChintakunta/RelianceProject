using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace Reliance
{
    public class FileHandler
    {
        private IWebBrowser webBrowser;


        public FileHandler(IWebBrowser webBrowser)
        {
            this.webBrowser = webBrowser;
        }
        public void Open(List<string> paths)
        {
            foreach (string path in paths)
            {
                Process.Start(path);
            }
        }

        public void Close(string process)
        {
            Array.ForEach(Process.GetProcessesByName(process), x => x.Kill());
        }

        private void ShowWindowDialog(string result)
        {

            _ = MessageBox.Show(result, "Result");
        }
    }
}
