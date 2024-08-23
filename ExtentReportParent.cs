using AventStack.ExtentReports;
using AventStack.ExtentReports.MarkupUtils;
using AventStack.ExtentReports.Reporter;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Sample;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using static Microsoft.Dynamics365.UIAutomation.Api.Pages.ActivityFeed;
using Status = AventStack.ExtentReports.Status;
using TestContext = Microsoft.VisualStudio.TestTools.UnitTesting.TestContext;


namespace Microsoft.Dynamics365.UIAutomation.CE.Reports
{
    [TestClass]
    public class ExtentReportParent
    {
        public static ExtentReports Extent;
        public static ExtentTest TestParent;
        public static ExtentTest Test;
        public string ScreenshotPath;
        public string Base64Code;
        public static string AssemblyName;
        public static int CaseCount;
        public TestContext TestContext { get; set; }


        [AssemblyInitialize] //executes only once at start
        public static void AssemblyInitialize(TestContext context)
        {
            CaseCount = 0;

            AssemblyName = Assembly.GetExecutingAssembly().GetName().Name;
            // The result is added under TestResult folder under EasyReproD365 folder

            var dir = context.TestDir + "\\";
            const string fileName = "ExtentReport.html";
            var htmlReporter = new ExtentSparkReporter(dir + fileName);

            // Add any additional contextual information
            Extent = new ExtentReports();
            Extent.AddSystemInfo("Browser", Enum.GetName(typeof(BrowserType), TestSettings.Options.BrowserType));
            Extent.AddSystemInfo("Test User", System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"]);
            Extent.AddSystemInfo("D365 CE Instance", System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"]);
            Extent.AttachReporter(htmlReporter);
            context.AddResultFile(fileName);
        }

        [TestInitialize] //executes before each test
        public void TestInitialize()
        {
            CaseCount++;
            // Get unit test description attribute
            var type = Type.GetType($"{TestContext.FullyQualifiedTestClassName}, {AssemblyName}");
            var methodInfo = type?.GetMethod(TestContext.TestName);
            var customAttributes = methodInfo?.GetCustomAttributes(false);
            DescriptionAttribute desc = null;
            if (customAttributes != null)
            {
                foreach (var n in customAttributes)
                {
                    desc = n as DescriptionAttribute;
                    if (desc != null)
                        break;
                }
            }

            //Create individual test under the relevant class
            Test = Extent.CreateTest(TestContext.TestName).AssignCategory(TestContext.FullyQualifiedTestClassName);
        }


        [TestCleanup] //executes after each test
        public void TestCleanup()
        {
            if (Test.Status == Status.Error)
                return;

            switch (TestContext.CurrentTestOutcome)
            {
                case UnitTestOutcome.Error:
                    Test.Fail("Test Failed - System Error");
                    break;
                case UnitTestOutcome.Passed:
                    Test.Pass("Test Passed");
                    break;
                case UnitTestOutcome.Failed:
                    Test.Fail("Test Failed");
                    break;
                case UnitTestOutcome.Inconclusive:
                    Test.Fail("Test Failed - Inconclusive");
                    break;
                case UnitTestOutcome.Timeout:
                    Test.Fail("Test Failed - Timeout");
                    break;
                case UnitTestOutcome.Aborted:
                    Test.Skip("Test Failed - Aborted / Not Runnable");
                    break;
                case UnitTestOutcome.InProgress:
                case UnitTestOutcome.Unknown:
                default:
                    Test.Fail("Test Failed - Unknown");
                    break;
            }


            ScreenshotPath = TestContext.TestDir + "\\In\\" + TestContext.TestName + CaseCount + ".png";
            Test.AddScreenCaptureFromPath(ScreenshotPath, TestContext.TestName + CaseCount);
        }


        [AssemblyCleanup] //executes only once at end of execution
        public static void AssemblyCleanup()
        {
            Extent.Flush();
        }

        public void TakeScreenshot(WebClient client, string title)
        {
            string path = TestContext.TestDir + "\\In\\" + title + ".png";
            client.Browser.Driver.WaitForTransaction();
            Screenshot screenshot = ((ITakesScreenshot)client.Browser.Driver).GetScreenshot();
            screenshot.SaveAsFile(path);
        }

        public void markTestcaseSatus(Exception e)
        {
            // Formats the exception details to look nice
            var message = e.Message + Environment.NewLine + e.StackTrace.Trim();
            var markup = MarkupHelper.CreateCodeBlock(message);
            Test.Fail(markup);
            throw e;
        }

        public void getErrorLogs(WebClient client)
        {
            client.Browser.Driver.SwitchTo().ActiveElement();
            if (client.Browser.Driver.FindElements(By.Id("dialogErrorViewlink")).Count > 0)
            {
                client.Browser.Driver.ClickWhenAvailable(By.Id("dialogErrorViewlink"));
                client.Browser.ThinkTime(2000);
                client.Browser.Driver.ClickWhenAvailable(By.Id("dialogErrorTexttext"));
                client.Browser.ThinkTime(4000);
                renameLatestFile();
            }
        }

        public void createNewFolder()
        {
            DateTime currentTime = DateTime.Now;
            string folderName = currentTime.ToString("yyyy-MM-dd_HH-mm-ss");
            string folderPath = Path.Combine(Directory.GetCurrentDirectory(), folderName);
            Directory.CreateDirectory(folderPath);
        }

        public void renameLatestFile()
        {
            string path = new BrowserOptions().DownloadsPath;
            DirectoryInfo directory = new DirectoryInfo(path);

            FileInfo latestFile = directory.GetFiles()
                                           .OrderByDescending(f => f.CreationTime)
                                           .FirstOrDefault();

            if (latestFile != null)
            {
                string newFileName = Path.Combine(directory.FullName, TestContext.TestName + "_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".txt");
                latestFile.MoveTo(newFileName);

                Console.WriteLine($"Renamed and changed extension of {latestFile.FullName} to {newFileName}");
            }
            else
            {
                Console.WriteLine("No files found in the folder");
            }

        }
    }
}
