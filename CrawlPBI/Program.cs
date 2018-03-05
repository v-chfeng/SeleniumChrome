using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.IE;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using CrawlPBI.Mode;
using Newtonsoft.Json;

namespace CrawlPBI
{
    class Program
    {
        static void Main(string[] args)
        {
            var pbiUrl = @"https://msit.powerbi.com/groups/me/settings/datasets";
            //var pbiUrl = "https://awsui.partners.extranet.microsoft.com/MOPET/?ProposalID=test";//@"https://msit.powerbi.com/groups/me/settings/datasets";

            IWebDriver webDriver = new ChromeDriver(GetChromeDriverService());//new PhantomJSDriver(GetPhantomJSDriverService());
            //IWebDriver webDriver = new PhantomJSDriver(GetPhantomJSDriverService());
            
            //IWebDriver webDriver = new InternetExplorerDriver(GetIEDriverService());
            webDriver.Navigate().GoToUrl(pbiUrl);

            webDriver.Manage().Timeouts().ImplicitWait = new TimeSpan(0, 1, 0);
            webDriver.Manage().Timeouts().PageLoad = new TimeSpan(0, 1, 0);

            // 登录microsoft账号
            if (webDriver.Url.Contains("https://login.microsoftonline.com"))
            {
                var accountInput = webDriver.FindElement(By.CssSelector("input[type='email']"));
                accountInput.SendKeys("");

                var loginEle = webDriver.FindElement(By.CssSelector("input[type='submit']"));
                loginEle.Click();

                WebDriverWait waitForSignInPage = new WebDriverWait(webDriver, new TimeSpan(0, 1, 0));
                waitForSignInPage.Until(d => d != null && d.Url != null && d.Url.Contains("https://msft.sts.microsoft.com"));

                if (webDriver.Url.Contains("https://msft.sts.microsoft.com"))
                {
                    var loginPasswordTextEle = webDriver.FindElement(By.CssSelector("a[class='actionLink']"));
                    //Screenshot shot = ((ITakesScreenshot)webDriver).GetScreenshot();
                    //shot.SaveAsFile(@"shot1.png");
                    Thread.Sleep(1000);
                    loginPasswordTextEle.Click();

                    var passwordInputEle = webDriver.FindElement(By.CssSelector("#passwordInput"));
                    passwordInputEle.SendKeys("");

                    var signInEle = webDriver.FindElement(By.CssSelector("#submitButton"));
                    signInEle.Click();
                }
            }

            (new WebDriverWait(webDriver, new TimeSpan(0, 1, 0))).Until(d => d!=null && d.Url != null && d.Url == pbiUrl);  

            Console.WriteLine("before find elements");
            var elements = webDriver.FindElements(By.XPath("//div[@class='collectionList']/div/ul/li"));
            //Screenshot shot2 = ((ITakesScreenshot)webDriver).GetScreenshot();
            //shot2.SaveAsFile(@"shot2.png");
            string[] reports = new string[] { "AccountReport", "Skill&AnalyticsReports", "xics_ContentReports", "xics_ConversationReports", "xics_GeneralReports", "xics_AppsReports" };
            List<RefreshResult> resultList = new List<RefreshResult>();
            foreach (var item in elements)
            {   
                // 找到一个期望的report
                if (reports.Contains(item.Text.Trim()))
                {  
                    Console.WriteLine($"find {item.Text}");
                    item.Click();
                    Thread.Sleep(1000);
                    var refreshHistoryEle = webDriver.FindElement(By.CssSelector(".refreshHistory"));
                    refreshHistoryEle.Click();

                    // 等待刷新记录弹窗
                    var refreshDialogEle = webDriver.FindElement(By.CssSelector(".dialog"));
                    var closeBtn = webDriver.FindElement(By.CssSelector(".dialog .confirm"));
                    RefreshResult result = GetFirstRefreshRecord(refreshDialogEle);
                    if (result != null)
                    {
                        result.Name = item.Text.Trim();
                        resultList.Add(result);
                    }

                    closeBtn.Click(); 
                }

            }

            bool isSuccess = true;
            //Console.WriteLine(JsonConvert.SerializeObject(resultList));
            foreach (var item in resultList)
            {
                string text = string.Format("{0:#################}\t{1}\t{2}", item.Name, item.Status, item.Date);
                Console.WriteLine(text);
                if (DateTime.Parse(item.Date).Day != DateTime.Now.Day)
                {
                    isSuccess = false;
                }
            }

            if (resultList.Count != reports.Length)
            {
                isSuccess = false;
            }

            Console.WriteLine($"Refresh success:{isSuccess}");
            //Console.WriteLine(webDriver.PageSource);
            //File.WriteAllText(@".\pbi.html", webDriver.PageSource);

            webDriver.Close();

            TimeSpan span = new TimeSpan(1, 0, 0);
            Thread.Sleep(span);
            webDriver.Quit();
        }

        private static RefreshResult GetFirstRefreshRecord(IWebElement dialogEle)
        {
            var firstTableBodyEle = dialogEle.FindElement(By.CssSelector(".historyTableBody"));
            if (firstTableBodyEle != null)
            {
                string startDate = firstTableBodyEle.FindElement(By.CssSelector(".historyStart")).Text;
                string Status = firstTableBodyEle.FindElement(By.CssSelector(".historyStatus")).Text;
                return new RefreshResult { Date = startDate, Status = Status };
            }

            return null; 
        }

        private static void CheckAllReportsStatus()
        {


        }

        private static void GetRefreshStatus()
        {
            
        }

        private static PhantomJSDriverService GetPhantomJSDriverService()
        {
            PhantomJSDriverService jsService = PhantomJSDriverService.CreateDefaultService();

            PhantomJSOptions options = new PhantomJSOptions();
            //jsService.agen
            //jsService.
            //jsService.Proxy;
            //jsService.ProxyAuthentication;

            return jsService;
        }

        private static ChromeDriverService GetChromeDriverService()
        {
            ChromeDriverService chromeService = ChromeDriverService.CreateDefaultService();


            return chromeService;
        }


        private static EdgeDriverService GetEdgeDriverService()
        {
            EdgeDriverService edgeService = EdgeDriverService.CreateDefaultService();

            return edgeService;
        }

        private static InternetExplorerDriverService GetIEDriverService()
        {
            string path = @"C:\Program Files\internet explorer\";
            string name = "iexplore.exe";
            InternetExplorerDriverService ieService = InternetExplorerDriverService.CreateDefaultService();

            //ieService.

            return ieService;
        }
    }
}
