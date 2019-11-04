using OpenQA.Selenium;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Reliance
{
    public interface IWebBrowser
    {

        string downloadPath { get; }
        void Invoke(string url);

        string GetWindowTitle();

        IWebElement WaitForElementVisible(string element);

        IWebElement WaitForElementClickable(string element);

        IWebElement WaitForElementClickable(IWebElement element);

        IWebElement FindElement(string element);

        void MovetoElement(string element);

        void MovetoElement(IWebElement element);

        void EnterText(string element, string input);

        void Click(string element);

        void ClickWithActions(string element);

        void ClickWithJavaScript(IWebElement element);

        List<IWebElement> FindElements(string element);

        IWebElement FindChildElement(string parentElement, int index, string childElement);

        object ExecuteScript(string script);

        void WaitForAllElementsToLoadInAPage();

        void Dispose();

        bool IsElementDisplayed(string element);

        bool IsElementEnabled(string element);

        void RefreshPage();

        void Pause(int timeOut);

        void PerformEnterAction();
    }
}
