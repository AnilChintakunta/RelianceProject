using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading;

namespace Reliance
{
    class EmailHandler : IEmailHandler
    {
        private IWebBrowser webBrowser;

        public EmailHandler(IWebBrowser webBrowser)
        {
            this.webBrowser = webBrowser;
        }
        public void LoginAndNavigateToInbox(string Username, string Password)
        {
            webBrowser.Invoke("https://accounts.google.com/signin/v2/identifier?continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&service=mail&sacu=1&rip=1&flowName=GlifWebSignIn&flowEntry=ServiceLogin");
            if (webBrowser.IsElementDisplayed(WebElements.TxtBoxEmail))
            {
                webBrowser.EnterText(WebElements.TxtBoxEmail, Username);
                webBrowser.Click(WebElements.BtnNext);
                webBrowser.WaitForElementVisible(WebElements.TxtBoxPassword);
                webBrowser.EnterText(WebElements.TxtBoxPassword, Password);
                webBrowser.Click(WebElements.BtnNext);
                webBrowser.WaitForElementVisible(WebElements.BtnInbox);

            }
            else if (webBrowser.IsElementDisplayed(WebElements.TxtBoxPassword))
            {
                webBrowser.EnterText(WebElements.TxtBoxPassword, Password);
                webBrowser.Click(WebElements.BtnNext);
                webBrowser.WaitForElementVisible(WebElements.BtnInbox);
            }
            else
            {
                webBrowser.WaitForElementVisible(WebElements.BtnInbox);
            }

        }

        public List<IWebElement> GetUnreadEmails()
        {
            List<IWebElement> result = new List<IWebElement>();
            int retry = 0;
            while (retry < 4)
            {
                try
                {
                    result = webBrowser.FindElements(WebElements.ElementUnreadEmails);
                    if (result.Count >= 1)
                    {
                        return result;
                    }
                    else
                    {
                        webBrowser.Pause(3);
                    }
                }
                catch (Exception e)
                {
                    webBrowser.WaitForAllElementsToLoadInAPage();
                    Console.WriteLine(e.Message);
                }
                retry++;
            }
            return result;
        }

        public Dictionary<string, string> DownloadAndGetFilePaths()
        {
            Dictionary<string, string> locations = new Dictionary<string, string>();
            webBrowser.WaitForElementClickable(WebElements.ElementIconDownloadAttachment);
            var attachments = webBrowser.FindElements(WebElements.ElementIconDownloadAttachment);
            foreach (var item in attachments)
            {
                string label = item.GetAttribute("aria-label");
                string fileName = label.Substring(("Download attachment ").Length);
                string fullPath = Path.Combine(webBrowser.downloadPath, fileName);
                item.SendKeys(Keys.Enter);
                var isDownloaded = false;
                while (!isDownloaded)
                {
                    isDownloaded = File.Exists(fullPath);
                    Thread.Sleep(1000);
                }
                locations.Add(fileName, fullPath);

            }
            return locations;
        }

        public void ReplyEmail(string text)
        {
            if (webBrowser.IsElementDisplayed(WebElements.BtnReply))
            {
                webBrowser.EnterText(WebElements.TxtBoxReplyMessage, text);
                webBrowser.Click(WebElements.BtnReply);
            }
        }

    }
}
