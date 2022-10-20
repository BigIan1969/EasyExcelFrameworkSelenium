using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using static EasyExcelFrameworkSelenium.EasyExcelFrameworkSelenium;

namespace EasyExcelFrameworkSelenium
{
    internal class Wait
    {
        private IWebDriver driver;

        public Wait(IWebDriver _driver)
        {
            driver = _driver;
        }
        public int timeout(string suppliedstring)
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
        public IWebElement bywait(Locator lc, bywaitconditions condition, int seconds = 60)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(1));
            DateTime endtime = DateTime.Now.AddSeconds(seconds);
            while (DateTime.Now < endtime)
            {
                IWebElement element;
                foreach (By by in lc.bys)
                {

                    switch (condition)
                    {
                        case bywaitconditions.VISIBLE:
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
                        case bywaitconditions.EXISTS:
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

        public IWebElement elewait(Locator lc, elewaitconditions condition, int seconds = 60)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(1));
            DateTime endtime = DateTime.Now.AddSeconds(seconds);
            while (DateTime.Now < endtime)
            {
                IWebElement element;
                foreach (IWebElement ele in lc.targets)
                {

                    switch (condition)
                    {
                        case elewaitconditions.CLICKABLE:
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
        public void waitalert(int seconds = 60)
        {
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.AlertIsPresent());
        }

    }

}

