using EasyExcelFramework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System.Data;

namespace EasyExcelFrameworkSelenium
{
    public class EasyExcelFrameworkSelenium
    {
        public EasyExcelF EasyExcel;
        public IWebDriver driver;
        private Wait waitclass;

        public enum elewaitconditions { CLICKABLE }
        public enum bywaitconditions { VISIBLE, EXISTS }

        public EasyExcelFrameworkSelenium(string filename)
        {
            //Implement Constructor for Framework Extension
            this.EasyExcel = new EasyExcelF(filename);
            init();
        }

        private void init()
        {
            EasyExcel.RegisterScreenShot(Screenshot);

            EasyExcel.RegisterMethod("LAUNCH", launchbrowser);
            
            EasyExcel.RegisterMethod("URL", gotourl);
            
            EasyExcel.RegisterMethod("GET BROWSER CAPABILITIES", getbrowsercapabilities);
            EasyExcel.RegisterMethod("GET URL", geturl);
            EasyExcel.RegisterMethod("GET ATTRIBUTE", getattribute);
            EasyExcel.RegisterMethod("IS SELECTED", isselected);
            EasyExcel.RegisterMethod("GET SELECTION", getselected);

            EasyExcel.RegisterMethod("ACCEPT ALERT", acceptalert);
            EasyExcel.RegisterMethod("DISMISS ALERT", dismisalert);

            EasyExcel.RegisterMethod("CLICK", click);
            EasyExcel.RegisterMethod("TYPE", type);
            EasyExcel.RegisterMethod("CLEAR", clear);
            EasyExcel.RegisterMethod("SUBMIT", submit);
            EasyExcel.RegisterMethod("CHECK", check);
            EasyExcel.RegisterMethod("UNCHECK", uncheck);
            EasyExcel.RegisterMethod("SELECT BY TEXT", selectbytext);
            EasyExcel.RegisterMethod("SELECT BY VALUE", selectbyvalue);
            EasyExcel.RegisterMethod("SELECT BY INDEX", selectbyindex);
            EasyExcel.RegisterMethod("DESELECT BY TEXT", deselectbytext);
            EasyExcel.RegisterMethod("DESELECT BY VALUE", deselectbyvalue);
            EasyExcel.RegisterMethod("DESELECT BY INDEX", deselectbyindex);
            EasyExcel.RegisterMethod("SELECT ALL", selectall);
            EasyExcel.RegisterMethod("DESELECT ALL", deselectall);

            EasyExcel.RegisterMethod("WAIT UNTIL ALERT IS PRESENT", waituntilalertpresent);
            EasyExcel.RegisterMethod("WAIT UNTIL ELEMENT EXISTS", waituntilcontrolexists);
            EasyExcel.RegisterMethod("WAIT UNTIL ELEMENT VISIBLE", waituntilcontrolvisible);
            EasyExcel.RegisterMethod("WAIT UNTIL ELEMENT CLICKABLE", waituntilcontrolclickable);

        }

        private bool launchbrowser(EasyExcelF ee, string[] parms)
        {
            string targetbrowser;
            int timeout=60;
            if (ee.Globals.ContainsKey("__EEF__TIMEOUT"))
            {
                timeout = (int)ee.Globals["__EEF__TIMEOUT"];
            }
            try
            {
                targetbrowser = ee.Interpreter.EvalToString(ee, parms[0], parms);
            }
            catch
            {
                targetbrowser = parms[0];
            }
            switch (targetbrowser.ToUpper())
            {
                case "CHROME":
                    ChromeOptions chromeopts = new ChromeOptions();

                    foreach (string opt in parms[1..])
                    {
                        if (!string.IsNullOrEmpty(opt))
                            chromeopts.AddArgument(opt);
                    }
                    driver = new ChromeDriver(ChromeDriverService.CreateDefaultService(), chromeopts, TimeSpan.FromMinutes(3));
                    _ = driver.Manage().Timeouts().PageLoad.Add(TimeSpan.FromSeconds(timeout));
                    break;
                case "CHROME HEADLESS":
                    ChromeOptions headlesschromeopts = new ChromeOptions();
                    headlesschromeopts.AddArgument("--headless");
                    headlesschromeopts.AddArgument("--disable-gpu");

                    foreach (string opt in parms[1..])
                    {
                        if (!string.IsNullOrEmpty(opt))
                            headlesschromeopts.AddArgument(opt);
                    }
                    driver = new ChromeDriver(ChromeDriverService.CreateDefaultService(), headlesschromeopts, TimeSpan.FromMinutes(3));
                    _ = driver.Manage().Timeouts().PageLoad.Add(TimeSpan.FromSeconds(timeout));
                    break;
                case "EDGE":
                    driver = new EdgeDriver();
                    _ = driver.Manage().Timeouts().PageLoad.Add(TimeSpan.FromSeconds(timeout)); 
                    break;
                case "FIREFOX":
                    driver = new FirefoxDriver();
                    _ = driver.Manage().Timeouts().PageLoad.Add(TimeSpan.FromSeconds(timeout));
                    break;

            }
            waitclass = new Wait(driver);
            return true;
        }
        private bool getbrowsercapabilities(EasyExcelF ee, string[] parms)
        {
            ICapabilities cap = ((ChromeDriver)driver).Capabilities;
            ee.Locals[parms[1]] = cap.GetCapability(parms[0]);
            return true;
        }
        private bool gotourl(EasyExcelF ee, string[] parms)
        {
            EasyExcelFramework.InterpreterClass interp = new EasyExcelFramework.InterpreterClass();
            string targ;
            try
            {
                targ = ee.Interpreter.EvalToString(ee, parms[0], parms);
            }
            catch
            {
                targ = parms[0];
            }
            driver.Navigate().GoToUrl(targ);
            return true;
        }
        private bool geturl(EasyExcelF ee, string[] parms)
        {
            ee.Locals[parms[0]] = driver.Url.ToString();
            return true;
        }
        private bool click(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.elewait(lc, elewaitconditions.CLICKABLE);
            if (ele != null)
            {
                ele.Click();
            }
            else
            {
                throw new Exception("Unable to find element: " + parms[0]);
            }
            return true;
        }
        private bool submit(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.elewait(lc, elewaitconditions.CLICKABLE);
            if (ele != null)
            {
                ele.Submit();
            }
            else
            {
                throw new Exception("Unable to find element: " + parms[0]);
            }
            return true;
        }
        private bool type(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.elewait(lc, elewaitconditions.CLICKABLE);
            if (ele != null)
            {
                string text;
                try
                {
                    text = ee.Interpreter.EvalToString(ee, parms[1], parms);
                }
                catch
                {
                    text = parms[1];
                }
                ele.SendKeys(text);
            }
            else
            {
                throw new Exception("Unable to find element: " + parms[0]);
            }
            return true;
        }
        private bool clear(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.elewait(lc, elewaitconditions.CLICKABLE);
            if (ele != null)
            {
                ele.Clear();
            }
            else
            {
                throw new Exception("Unable to find element: " + parms[0]);
            }
            return true;
        }
        private bool acceptalert(EasyExcelF ee, string[] parms)
        {
            waitclass.waitalert(waitclass.timeout(parms[0]));
            driver.SwitchTo().Alert().Accept();
            return true;
        }
        private bool dismisalert(EasyExcelF ee, string[] parms)
        {
            waitclass.waitalert(waitclass.timeout(parms[0]));
            driver.SwitchTo().Alert().Dismiss();
            return true;
        }
        private bool check(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.elewait(lc, elewaitconditions.CLICKABLE);
            if (ele.Selected)
                return true;
            else
                ele.Click();
            return true;
        }
        private bool uncheck(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.elewait(lc, elewaitconditions.CLICKABLE);
            if (!ele.Selected)
                return true;
            else
                ele.Click();
            return true;
        }
        private bool getattribute(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.bywait(lc, bywaitconditions.EXISTS);
            ee.Locals[parms[2]] = ele.GetAttribute(parms[1]);
            return true;
        }
        private bool isselected(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.bywait(lc, bywaitconditions.EXISTS);
            ee.Locals[parms[1]] = ele.Selected.ToString().ToLower();
            return true;
        }

        private bool selectbytext(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.bywait(lc, bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.SelectByText(parms[1]);
            return true;
        }
        private bool selectbyvalue(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.bywait(lc, bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.SelectByValue(parms[1]);
            return true;
        }
        private bool selectbyindex(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.bywait(lc, bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.SelectByIndex(int.Parse(parms[1]));
            return true;
        }
        private bool selectall(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.bywait(lc, bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            for (int i = 0; i < selectele.Options.Count; i++)
            {

                selectele.SelectByIndex(i);
            }
            return true;
        }
        private bool getselected(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.bywait(lc, bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            string result = "";
            for (int i = 0; i < selectele.Options.Count; i++)
            {
                if (selectele.Options[i].Selected)
                {
                    if (string.IsNullOrEmpty(result))
                    {
                        result = selectele.Options[i].Text.ToString();
                    }
                    else
                    {
                        result += "|" + selectele.Options[i].Text.ToString();
                    }
                }
            }
            ee.Locals[parms[1]] = result;
            return true;
        }
        private bool deselectall(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.bywait(lc, bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.DeselectAll();
            return true;
        }

        private bool deselectbytext(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.bywait(lc, bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.DeselectByText(parms[1]);
            return true;
        }
        private bool deselectbyvalue(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.bywait(lc, bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.DeselectByValue(parms[1]);
            return true;
        }
        private bool deselectbyindex(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.bywait(lc, bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.DeselectByIndex(int.Parse(parms[1]));
            return true;
        }
        private bool waituntilalertpresent(EasyExcelF ee, string[] parms)
        {
            waitclass.waitalert(waitclass.timeout(parms[0]));
            return true;
        }
        private bool waituntilcontrolexists(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            waitclass.bywait(lc, bywaitconditions.EXISTS, waitclass.timeout(parms[0]));
            return true;
        }
        private bool waituntilcontrolvisible(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            waitclass.bywait(lc, bywaitconditions.VISIBLE, waitclass.timeout(parms[0]));
            return true;
        }
        private bool waituntilcontrolclickable(EasyExcelF ee, string[] parms)
        {
            Locator lc = new Locator(ee, this.driver, parms[0]);
            waitclass.elewait(lc, elewaitconditions.CLICKABLE, waitclass.timeout(parms[0]));
            return true;
        }
        public class Locator
        {
            public string original;
            public List<IWebElement> targets;
            public List<By> bys;
            public Locator(EasyExcelF ee, IWebDriver driver, string loc)
            {
                original = loc;
                targets = new List<IWebElement>();
                List<string> targstrs = new List<string>();
                bys = new List<By>();
                if (ee.Worksheets.ContainsKey("@Locators"))
                {
                    string col1 = ee.Worksheets["@Locators"].Columns[0].ColumnName;
                    DataRow[] rows = ee.Worksheets["@Locators"].Select(col1 + " Like'" + loc + "'");
                    foreach (DataRow row in rows)
                    {
                        foreach (string col in row.ItemArray)
                        {
                            targstrs.Add(col.ToString());
                        }
                    }
                }
                else
                {
                    foreach (var str in loc.Split("|"))
                    {
                        targstrs.Add(str);
                    }
                }
                if (targstrs.Count == 0)
                    throw new IndexOutOfRangeException("Cannot have blank target");
                foreach (string targstr in targstrs)
                {
                    char[] c = { ':' };
                    string[] locs = targstr.Split(c, 2);
                    switch (locs[0].ToUpper())
                    {
                        case "ID":
                            bys.Add(By.Id(locs[1]));
                            foreach (IWebElement ele in driver.FindElements(By.Id(locs[1])))
                            {
                                targets.Add(ele);
                            }
                            break;
                        case "NAME":
                            bys.Add(By.Name(locs[1]));
                            foreach (IWebElement ele in driver.FindElements(By.Name(locs[1])))
                            {
                                targets.Add(ele);
                            }
                            break;
                        case "CLASSNAME":
                            bys.Add(By.ClassName(locs[1]));
                            foreach (IWebElement ele in driver.FindElements(By.ClassName(locs[1])))
                            {
                                targets.Add(ele);
                            }
                            break;
                        case "CSS":
                            bys.Add(By.CssSelector(locs[1]));
                            foreach (IWebElement ele in driver.FindElements(By.CssSelector(locs[1])))
                            {
                                targets.Add(ele);
                            }
                            break;
                        case "XPATH":
                            bys.Add(By.XPath(locs[1]));
                            foreach (IWebElement ele in driver.FindElements(By.XPath(locs[1])))
                            {
                                targets.Add(ele);
                            }
                            break;
                        case "LINKTEXT":
                        case "LINK":
                            bys.Add(By.LinkText(locs[1]));
                            foreach (IWebElement ele in driver.FindElements(By.LinkText(locs[1])))
                            {
                                targets.Add(ele);
                            }
                            break;
                        case "PARTIALLINKTEXT":
                        case "PARTIALLINK":
                            bys.Add(By.PartialLinkText(locs[1]));
                            foreach (IWebElement ele in driver.FindElements(By.PartialLinkText(locs[1])))
                            {
                                targets.Add(ele);
                            }
                            break;
                        case "LABEL":
                            bys.Add(By.XPath("//label[text()='" + locs[1] + "']/following-sibling::input"));
                            foreach (IWebElement ele in driver.FindElements(By.XPath("//label[text()='" + locs[1] + "']/following-sibling::input")))
                            {
                                targets.Add(ele);
                            }
                            break;
                        case "TEXT":
                            bys.Add(By.XPath("//*[text()='" + locs[1] + "']"));
                            foreach (IWebElement ele in driver.FindElements(By.XPath("//*[text()='" + locs[1] + "']")))
                            {
                                targets.Add(ele);
                            }
                            break;
                        case "TAGNAME":
                            bys.Add(By.TagName(locs[1]));
                            foreach (IWebElement ele in driver.FindElements(By.TagName(locs[1])))
                            {
                                targets.Add(ele);
                            }
                            break;
                        default:
                            if (locs[0].Substring(0, 1) == "/" | locs[0].Substring(0, 1) == "(")
                            {
                                bys.Add(By.XPath(locs[0]));
                                foreach (IWebElement ele in driver.FindElements(By.XPath(locs[0])))
                                {
                                    targets.Add(ele);
                                }
                            }
                            else
                            {
                                bys.Add(By.CssSelector(locs[0]));
                                foreach (IWebElement ele in driver.FindElements(By.CssSelector(locs[0])))
                                {
                                    targets.Add(ele);
                                }
                            }
                            break;
                    }
                }
            }
        }
        private string Screenshot(string defaultpath)
        {
            Screenshot TakeScreenshot = ((ITakesScreenshot)driver).GetScreenshot();
            Random randNo = new Random();
            string fname = "\\"+randNo.Next(0, 10000).ToString() + ".png";
            TakeScreenshot.SaveAsFile(defaultpath+fname);
            return defaultpath + fname;
        }
    }
}