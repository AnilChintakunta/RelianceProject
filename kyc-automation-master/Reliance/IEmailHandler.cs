using OpenQA.Selenium;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Reliance
{
    public interface IEmailHandler
    {
        void LoginAndNavigateToInbox(string Username, string Password);

        List<IWebElement> GetUnreadEmails();

        Dictionary<string, string> DownloadAndGetFilePaths();

        void ReplyEmail(string text);

    }
}
