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
        public IWebDriver? driver;
        private Wait? waitclass;

        public enum Elewaitconditions { CLICKABLE }
        public enum Bywaitconditions { VISIBLE, EXISTS }

        public EasyExcelFrameworkSelenium(string filename)
        {
            //Implement Constructor for Framework Extension
            this.EasyExcel = new EasyExcelF(filename);
            Init();
        }

        private void Init()
        {
            EasyExcel.RegisterScreenShot(Screenshot);

            EasyExcel.RegisterMethod("LAUNCH", Launchbrowser);
            
            EasyExcel.RegisterMethod("URL", Gotourl);
            
            EasyExcel.RegisterMethod("GET BROWSER CAPABILITIES", Getbrowsercapabilities);
            EasyExcel.RegisterMethod("GET URL", Geturl);
            EasyExcel.RegisterMethod("GET ATTRIBUTE", Getattribute);
            EasyExcel.RegisterMethod("IS SELECTED", Isselected);
            EasyExcel.RegisterMethod("GET SELECTION", Getselected);

            EasyExcel.RegisterMethod("ACCEPT ALERT", Acceptalert);
            EasyExcel.RegisterMethod("DISMISS ALERT", Dismisalert);

            EasyExcel.RegisterMethod("CLICK", Click);
            EasyExcel.RegisterMethod("TYPE", Type);
            EasyExcel.RegisterMethod("CLEAR", Clear);
            EasyExcel.RegisterMethod("SUBMIT", Submit);
            EasyExcel.RegisterMethod("CHECK", Check);
            EasyExcel.RegisterMethod("UNCHECK", Uncheck);
            EasyExcel.RegisterMethod("SELECT BY TEXT", Selectbytext);
            EasyExcel.RegisterMethod("SELECT BY VALUE", Selectbyvalue);
            EasyExcel.RegisterMethod("SELECT BY INDEX", Selectbyindex);
            EasyExcel.RegisterMethod("DESELECT BY TEXT", Deselectbytext);
            EasyExcel.RegisterMethod("DESELECT BY VALUE", Deselectbyvalue);
            EasyExcel.RegisterMethod("DESELECT BY INDEX", Deselectbyindex);
            EasyExcel.RegisterMethod("SELECT ALL", Selectall);
            EasyExcel.RegisterMethod("DESELECT ALL", Deselectall);

            EasyExcel.RegisterMethod("WAIT UNTIL ALERT IS PRESENT", Waituntilalertpresent);
            EasyExcel.RegisterMethod("WAIT UNTIL ELEMENT EXISTS", Waituntilcontrolexists);
            EasyExcel.RegisterMethod("WAIT UNTIL ELEMENT VISIBLE", Waituntilcontrolvisible);
            EasyExcel.RegisterMethod("WAIT UNTIL ELEMENT CLICKABLE", Waituntilcontrolclickable);

            EasyExcel.RegisterMethod("SWITCH FRAME DEFAULT", Switchframedefault);
            EasyExcel.RegisterMethod("SWITCH FRAME BY INDEX", Switchframebyname_id_index);
            EasyExcel.RegisterMethod("SWITCH FRAME BY NAME", Switchframebyname_id_index);
            EasyExcel.RegisterMethod("SWITCH FRAME BY ID", Switchframebyname_id_index);
            EasyExcel.RegisterMethod("SWITCH FRAME BY ELEMENT", Switchframebyelement);
            EasyExcel.RegisterMethod("SWITCH FRAME PARENT", Switchframeparent);
        }

        private bool Launchbrowser(EasyExcelF ee, string[] parms)
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
                    goto default;
                case "CHROME HEADLESS":
                    ChromeOptions headlesschromeopts = new();
                    headlesschromeopts.AddArgument("--headless");
                    headlesschromeopts.AddArgument("--disable-gpu");

                    foreach (string opt in parms[1..])
                    {
                        if (!string.IsNullOrEmpty(opt))
                            headlesschromeopts.AddArgument(opt);
                    }
                    driver = new ChromeDriver(ChromeDriverService.CreateDefaultService(), headlesschromeopts, TimeSpan.FromMinutes(3));
                    break;
                case "EDGE":
                    driver = new EdgeDriver();
                    break;
                case "FIREFOX":
                    driver = new FirefoxDriver();
                    break;
                default:
                    ChromeOptions chromeopts = new();

                    foreach (string opt in parms[1..])
                    {
                        if (!string.IsNullOrEmpty(opt))
                            chromeopts.AddArgument(opt);
                    }
                    driver = new ChromeDriver(ChromeDriverService.CreateDefaultService(), chromeopts, TimeSpan.FromMinutes(3));
                    break;

            }
            _ = driver.Manage().Timeouts().PageLoad.Add(TimeSpan.FromSeconds(timeout));
            waitclass = new Wait(driver);
            return true;
        }
        private bool Getbrowsercapabilities(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            ICapabilities cap = ((ChromeDriver)driver).Capabilities;
            ee.Locals[parms[1]] = cap.GetCapability(parms[0]);
            return true;
        }
        private bool Gotourl(EasyExcelF ee, string[] parms)
        {
            if (driver==null || waitclass==null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
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
        private bool Geturl(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            ee.Locals[parms[0]] = driver.Url.ToString();
            return true;
        }
        private bool Click(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.Elewait(lc, Elewaitconditions.CLICKABLE);
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
        private bool Submit(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.Elewait(lc, Elewaitconditions.CLICKABLE);
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
        private bool Type(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, driver, parms[0]);
            IWebElement ele = waitclass.Elewait(lc, Elewaitconditions.CLICKABLE);
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
        private bool Clear(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.Elewait(lc, Elewaitconditions.CLICKABLE);
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
        private bool Acceptalert(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            waitclass.Waitalert(Wait.Timeout(parms[0]));
            driver.SwitchTo().Alert().Accept();
            return true;
        }
        private bool Dismisalert(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            waitclass.Waitalert(Wait.Timeout(parms[0]));
            driver.SwitchTo().Alert().Dismiss();
            return true;
        }
        private bool Check(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.Elewait(lc, Elewaitconditions.CLICKABLE);
            if (ele.Selected)
                return true;
            else
                ele.Click();
            return true;
        }
        private bool Uncheck(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.Elewait(lc, Elewaitconditions.CLICKABLE);
            if (!ele.Selected)
                return true;
            else
                ele.Click();
            return true;
        }
        private bool Getattribute(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.Bywait(lc, Bywaitconditions.EXISTS);
            ee.Locals[parms[2]] = ele.GetAttribute(parms[1]);
            return true;
        }
        private bool Isselected(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.Bywait(lc, Bywaitconditions.EXISTS);
            ee.Locals[parms[1]] = ele.Selected.ToString().ToLower();
            return true;
        }

        private bool Selectbytext(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, driver, parms[0]);
            IWebElement ele = waitclass.Bywait(lc, Bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.SelectByText(parms[1]);
            return true;
        }
        private bool Selectbyvalue(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, driver, parms[0]);
            IWebElement ele = waitclass.Bywait(lc, Bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.SelectByValue(parms[1]);
            return true;
        }
        private bool Selectbyindex(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, driver, parms[0]);
            IWebElement ele = waitclass.Bywait(lc, Bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.SelectByIndex(int.Parse(parms[1]));
            return true;
        }
        private bool Selectall(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.Bywait(lc, Bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            for (int i = 0; i < selectele.Options.Count; i++)
            {

                selectele.SelectByIndex(i);
            }
            return true;
        }
        private bool Getselected(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.Bywait(lc, Bywaitconditions.EXISTS);
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
        private bool Deselectall(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.Bywait(lc, Bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.DeselectAll();
            return true;
        }

        private bool Deselectbytext(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.Bywait(lc, Bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.DeselectByText(parms[1]);
            return true;
        }
        private bool Deselectbyvalue(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.Bywait(lc, Bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.DeselectByValue(parms[1]);
            return true;
        }
        private bool Deselectbyindex(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            IWebElement ele = waitclass.Bywait(lc, Bywaitconditions.EXISTS);
            var selectele = new SelectElement(ele);
            selectele.DeselectByIndex(int.Parse(parms[1]));
            return true;
        }
        private bool Waituntilalertpresent(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            waitclass.Waitalert(Wait.Timeout(parms[0]));
            return true;
        }
        private bool Waituntilcontrolexists(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            waitclass.Bywait(lc, Bywaitconditions.EXISTS, Wait.Timeout(parms[0]));
            return true;
        }
        private bool Waituntilcontrolvisible(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            waitclass.Bywait(lc, Bywaitconditions.VISIBLE, Wait.Timeout(parms[0]));
            return true;
        }
        private bool Waituntilcontrolclickable(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            waitclass.Elewait(lc, Elewaitconditions.CLICKABLE, Wait.Timeout(parms[0]));
            return true;
        }
        private bool Switchframedefault(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            driver.SwitchTo().DefaultContent();
            return true;
        }
        private bool Switchframebyname_id_index(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            if (int.TryParse(parms[0], out _))
                driver.SwitchTo().Frame(int.Parse(parms[0]));
            else
                driver.SwitchTo().Frame(parms[0]);
            return true;
        }
        private bool Switchframebyelement(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            Locator lc = new(ee, this.driver, parms[0]);
            driver.SwitchTo().Frame(waitclass.Bywait(lc, Bywaitconditions.EXISTS));
            return true;
        }
        private bool Switchframeparent(EasyExcelF ee, string[] parms)
        {
            if (driver == null || waitclass == null)
            {
                throw new InvalidOperationException("Browser is in open");
            }
            driver.SwitchTo().ParentFrame();
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
                List<string> targstrs = new();
                bys = new List<By>();
                if (ee.Worksheets.ContainsKey("@Locators"))
                {
                    string col1 = ee.Worksheets["@Locators"].Columns[0].ColumnName;
                    DataRow[] rows = ee.Worksheets["@Locators"].Select(col1 + " Like'" + loc + "'");
                    foreach (DataRow row in rows)
                    {
                        foreach (string col in row.ItemArray.Cast<string>())
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
            Random randNo = new();
            string fname = "\\"+randNo.Next(0, 10000).ToString() + ".png";
            TakeScreenshot.SaveAsFile(defaultpath+fname);
            return defaultpath + fname;
        }
    }
}