// <copyright file="WebBrowser.cs" company="CAWStudios">
//     ChimpsAtWorkStudios. All rights reserved.
// </copyright>
// <author>RaviKiran</author>

namespace Reliance
{
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using OpenQA.Selenium.Firefox;
    using OpenQA.Selenium.Interactions;
    using OpenQA.Selenium.Support.UI;
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using ExpectedConditions = SeleniumExtras.WaitHelpers.ExpectedConditions;

    public class WebBrowser : IWebBrowser, IDisposable
    {
        private readonly IWebDriver driver;

        private readonly WebDriverWait wait;

        private readonly DefaultWait<IWebDriver> defaultWait;

        private readonly IJavaScriptExecutor jse;

        private readonly Actions actions;

        public string downloadPath => Path.GetFullPath("..\\..\\Files");

        public WebBrowser(string browserName)
        {
            switch (browserName.ToLower())
            {
                case "chrome":
                    {
                        ChromeOptions options = new ChromeOptions();
                        string defaultChromePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Google\\Chrome\\User Data";
                        options.AddUserProfilePreference("download.default_directory", downloadPath);
                        options.AddUserProfilePreference("download.prompt_for_download", false);
                        options.AddUserProfilePreference("disable-popup-blocking", "true");
                        options.AddArguments("--disable-notifications", "--disable-infobars", "--disable-gpu", "--autoplay-policy=no-user-gesture-required");
                        var driverService = ChromeDriverService.CreateDefaultService();
                        driverService.LogPath = @"../../chromedriver.log";
                        driver = new ChromeDriver(driverService, options);
                        driver.Manage().Window.Maximize();
                        break;
                    }

                case "firefox":
                    {
                        _ = new FirefoxOptions();
                        driver = new FirefoxDriver();
                        driver.Manage().Window.Maximize();
                        break;
                    }
            }

            wait = new WebDriverWait(this.driver, TimeSpan.FromSeconds(50000));
            defaultWait = new DefaultWait<IWebDriver>(driver);
            jse = (IJavaScriptExecutor)driver;
            actions = new Actions(driver);
        }

        public void Invoke(string url)
        {
            driver.Navigate().GoToUrl(url);
        }

        public string GetWindowTitle()
        {
            return driver.Title;
        }

        public IWebElement WaitForElementVisible(string element)
        {
            return wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(element)));
        }

        public IWebElement WaitForElementClickable(string element)
        {
            return wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(element)));
        }

        public IWebElement WaitForElementClickable(IWebElement element)
        {
            return wait.Until(ExpectedConditions.ElementToBeClickable(element));
        }

        public IWebElement FindElement(string element)
        {
            IWebElement result;

            try
            {
                result = driver.FindElement(By.CssSelector(element));
            }
            catch (Exception e)
            {
                defaultWait.Timeout = TimeSpan.FromSeconds(10);
                defaultWait.PollingInterval = TimeSpan.FromMilliseconds(250);
                defaultWait.IgnoreExceptionTypes(typeof(NoSuchElementException));
                result = defaultWait.Until(x => x.FindElement(By.CssSelector(element)));
                Console.WriteLine(e.Message);
            }
            return result;
        }

        public void MovetoElement(IWebElement element)
        {
            Actions actions = new Actions(driver);
            actions.MoveToElement(element).Perform();
        }

        public void EnterText(string element, string input)
        {
            FindElement(element).SendKeys(input);
        }

        public void Click(string element)
        {
            try
            {
                FindElement(element).Click();
            }
            catch (StaleElementReferenceException stle)
            {
                defaultWait.Timeout = TimeSpan.FromSeconds(40);
                defaultWait.PollingInterval = TimeSpan.FromMilliseconds(500);
                defaultWait.Until(ExpectedConditions.ElementToBeClickable(FindElement(element))).Click();
                Console.WriteLine(stle.Message);
            }
        }

        public void ClickWithActions(string element)
        {
            IWebElement ele;
            try
            {
               ele =  FindElement(element);
               actions.MoveToElement(ele).Click().Perform();
            }
            catch (Exception)
            {
                defaultWait.Timeout = TimeSpan.FromSeconds(40);
                defaultWait.PollingInterval = TimeSpan.FromMilliseconds(500);
                defaultWait.IgnoreExceptionTypes(typeof(NoSuchElementException));
                ele = FindElement(element);
                actions.MoveToElement(ele).Click().Perform();
            }
        }

        public void ClickWithJavaScript(IWebElement element)
        {
            jse.ExecuteScript("window.scrollTo(0," + element.Location.Y + ")");
            jse.ExecuteScript("arguments[0].click();", element);
        }

        public List<IWebElement> FindElements(string element)
        {
            List<IWebElement> result = driver.FindElements(By.CssSelector(element)).ToList();
            if (result.Count >= 1)
            {
                return result;
            }
            else
            {
                try
                {
                    defaultWait.Timeout = TimeSpan.FromSeconds(10);
                    defaultWait.PollingInterval = TimeSpan.FromMilliseconds(250);
                    defaultWait.IgnoreExceptionTypes(typeof(NoSuchElementException));
                    result = defaultWait.Until(x => x.FindElements(By.CssSelector(element))).ToList();
                }
                catch (Exception)
                {
                    
                }
                return result;
            }
        }

        public IWebElement FindChildElement(string parentElement, int index, string childElement)
        {
            return driver.FindElements(By.CssSelector(parentElement))[index].FindElement(By.CssSelector(childElement));
        }

        public bool IsElementDisplayed(string element)
        {
            int attempts = 0;
            while (attempts < 3)
            {
                try
                {
                    List<IWebElement> elements = FindElements(element);
                    if (elements.Count > 0)
                    {
                        return elements.ElementAt(0).Displayed && elements.ElementAt(0).Enabled;
                    }
                    else
                    {
                        return false;
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine($"Element: {element} is not displayed");
                }
                attempts++;
            }
            return false;
        }

        public bool IsElementEnabled(string element)
        {
            int attempts = 0;
            while (attempts < 3)
            {
                try
                {
                    IReadOnlyCollection<IWebElement> elements = this.FindElements(element);
                    if (elements.Count > 0)
                    {
                        return elements.ElementAt(0).Enabled;
                    }
                    return false;
                }
                catch (Exception)
                {
                    Console.WriteLine($"Element: {element} is not displayed");
                }
                attempts++;
            }
            return false;
        }

        public bool IsClickable(string element)
        {
            try
            {
                wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(element)));
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        public void RefreshPage()
        {
            this.driver.Navigate().Refresh();
        }

        public object ExecuteScript(string script)
        {
            try
            {
                return this.jse.ExecuteScript(script);
            }
            catch (Exception)
            {
            }
            return new object();
        }

        public void Pause(int timeOut)
        {
            Thread.Sleep(TimeSpan.FromSeconds(timeOut));
        }

        public void WaitForAllElementsToLoadInAPage()
        {
            try
            {
                bool state = false;
                var pageLoadCheck = this.driver as IJavaScriptExecutor;
                while (!state)
                {
                    object obj = this.jse.ExecuteScript("return (document.readyState == 'complete' && navigator.onLine == true)");
                    state = Convert.ToBoolean(obj);
                    defaultWait.Timeout = TimeSpan.FromSeconds(40);
                    defaultWait.PollingInterval = TimeSpan.FromMilliseconds(500);
                    defaultWait.IgnoreExceptionTypes(typeof(NoSuchElementException));
                }
            }
            catch (Exception)
            {
            }
        }

        public void Dispose()
        {
            this.driver.Close();
            this.driver.Dispose();
            this.driver.Quit();
        }

        public void DeleteCookies()
        {
            this.driver.Manage().Cookies.DeleteAllCookies();
            Thread.Sleep(5000);
        }

        public void KillAllBrowserProcesses(string browser)
        {
            if (browser == "chrome")
            {
                Process[] chromeProcess = Process.GetProcessesByName("chromedriver");
                foreach (var process in chromeProcess)
                {
                    process.Kill();
                }
            }
            else
            {
                Process[] firefoxProcess = Process.GetProcessesByName("firefoxdriver");
                foreach (var process in firefoxProcess)
                {
                    process.Kill();
                }
            }
        }

        public void MovetoElement(string element)
        {
            IWebElement ele = FindElement(element);
            actions.MoveToElement(ele);
        }

        public void PerformEnterAction()
        {
            actions.SendKeys(Keys.Enter).Perform();
        }
    }
}