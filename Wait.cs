using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using static EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium;

namespace EasyExcelFrameworkSelenium
{
    internal class Wait
    {
        private readonly IWebDriver driver;

        public Wait(IWebDriver _driver)
        {
            driver = _driver;
        }
        public static int Timeout(string suppliedstring)
        {
            int timeout;
            try
            {
                timeout = int.Parse(suppliedstring);
            }
            catch
            {
                timeout = 60;
            }
            return (timeout);

        }
        public IWebElement? Bywait(Locator lc, Bywaitconditions condition, int seconds = 60)
        {
            WebDriverWait wait = new(driver, TimeSpan.FromMilliseconds(1));
            DateTime endtime = DateTime.Now.AddSeconds(seconds);
            while (DateTime.Now < endtime)
            {
                IWebElement? element;
                foreach (By by in lc.bys)
                {

                    switch (condition)
                    {
                        case Bywaitconditions.VISIBLE:
                            {
                                try
                                {
                                    element = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(by));
                                }
                                catch 
                                {
                                    element = null;
                                }
                                break;
                            }
                        case Bywaitconditions.EXISTS:
                            {
                                try
                                {
                                    element = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(by));
                                }
                                catch
                                {
                                    element = null;
                                }

                                break;
                            }


                        default:
                            {
                                element = null;
                                break;
                            }
                    }
                    if (element != null)
                    {
                        return element;

                    }
                }
            }
            return null;
        }

        public IWebElement? Elewait(Locator lc, Elewaitconditions condition, int seconds = 60)
        {
            WebDriverWait wait = new(driver, TimeSpan.FromMilliseconds(1));
            DateTime endtime = DateTime.Now.AddSeconds(seconds);
            while (DateTime.Now < endtime)
            {
                IWebElement? element;
                foreach (IWebElement ele in lc.targets)
                {

                    switch (condition)
                    {
                        case Elewaitconditions.CLICKABLE:
                            {
                                element = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(ele));
                                break;
                            }
                        default:
                            {
                                element = null;
                                break;
                            }
                    }
                    if (element != null)
                    {
                        if (element == ele)
                        {
                            return element;
                        }
                    }
                }
            }
            return null;
        }
        public void Waitalert(int seconds = 60)
        {
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(seconds));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.AlertIsPresent());
        }

    }

}

