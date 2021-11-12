using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using RelevantCodes.ExtentReports;
using System;
using System.Configuration;
using System.IO;

namespace SeleniumCSharpMSTest.GeneralFunctions
{
    public class BaseClass : Helper
    {

        public static string uRL = ConfigurationManager.AppSettings["URL"];
        public static string ESTuRL = ConfigurationManager.AppSettings["ESTURL"];

        public static IWebDriver driver;
        public static ExtentReports extentReport;
        public ExtentTest extentTest;

        public static string testr = null;

        public string testName = "";
        public string testDataIteration = "";
        public static string projectPath = SetProjectPath();
        public static string reportPath = SetReport();

        /// <summary>
        /// Method to Start the Drivers - This should be called in [TestInitialize] function
        /// </summary>
        public void TestInitialize(string browser = "Config")
        {
            extentReport = new ExtentReports(reportPath, false);
            extentReport.AddSystemInfo("Host Name", "Testhouse")
            .AddSystemInfo("Environment", "QA")
            .AddSystemInfo("User Name", "Testhouse");
            extentReport.LoadConfig(projectPath + "\\Reports\\config.xml");
            testName = TestContext.TestName.ToString();

            try
            {
                testDataIteration = TestContext.DataRow.Table.Rows.IndexOf(TestContext.DataRow).ToString();
            }
            catch (Exception)
            {

            }
            if (testDataIteration == "")
            {
                testDataIteration = "0";
            }
            extentTest = extentReport.StartTest(testName + " : Iteration" + testDataIteration);
            if (browser == "Config" || browser == null || browser == "")
            {
                DriverSetup();
            }
            else
            {
                DriverSetup(browser);
            }
        }

        /// <summary>
        /// Method to initiate the Driver Setup
        /// </summary>
        public void DriverSetup(string browser = null)
        {
            string filePath = projectPath + @"\Drivers";
            if (browser == null)
            {
                browser = ConfigurationManager.AppSettings["browser"];
            }
            if (string.Equals(browser, "IE", StringComparison.OrdinalIgnoreCase) || string.Equals(browser, "internet explorer", StringComparison.OrdinalIgnoreCase))
            {
                driver = new InternetExplorerDriver(filePath);
            }
            else if (browser == "Firefox")
            {
                FirefoxDriverService service = FirefoxDriverService.CreateDefaultService(filePath);
                service.FirefoxBinaryPath = "";
                if (File.Exists(@"C:\Program Files\Mozilla Firefox\firefox.exe"))
                {
                    service.FirefoxBinaryPath = @"C:\Program Files\Mozilla Firefox\firefox.exe";
                }
                else
                {
                    service.FirefoxBinaryPath = @"C:\Program Files (x86)\Mozilla Firefox\firefox.exe";
                }

                driver = new FirefoxDriver(service);
            }
            else if (browser == "FirefoxOld")
            {
                driver = new FirefoxDriver();
            }
            else if (browser == "Edge")
            {
                driver = new EdgeDriver(filePath);
            }
            else if (browser == "Chrome")
            {

                var option = new ChromeOptions();
                option.AddArgument("no-sandbox");
                driver = new ChromeDriver(filePath, option, TimeSpan.FromMinutes(3));
            }
        }

        /// <summary>
        /// Method to Clean up the Drivers and Write the report - This should be called in [TestCleanup] function
        /// </summary>
        public void MyTestCleanup(int testDataRowCount = 0, string createAndUploadZipFile = "No", string uploadScreenshots = "No")
        {

            try
            {
                driver.Close();
                driver.Quit();
                extentReport.EndTest(extentTest);
                extentReport.Flush();
                extentReport.Close();
                HideLogoVersion(reportPath);
                if (testDataRowCount != 0)
                {
                    testDataRowCount = testDataRowCount - 1;
                }
                if (testDataIteration == testDataRowCount.ToString())
                {
                    if (uploadScreenshots != "No")
                    {
                        UploadScreenshots(testContext);
                        AttachFilesToLog(reportPath, testContext);
                    }

                    if (createAndUploadZipFile != "No")
                    {
                        AddReportFileToZip(reportPath);
                        AttachFilesToLog(zipPath, testContext);
                    }
                }
            }
            catch (Exception)
            {
                extentReport.EndTest(extentTest);
                extentReport.Flush();
                extentReport.Close();
                HideLogoVersion(reportPath);
                if (testDataRowCount != 0)
                {
                    testDataRowCount = testDataRowCount - 1;
                }
                if (testDataIteration == testDataRowCount.ToString())
                {
                    if (uploadScreenshots != "No")
                    {
                        UploadScreenshots(testContext);
                        AttachFilesToLog(reportPath, testContext);
                    }

                    if (createAndUploadZipFile != "No")
                    {
                        AddReportFileToZip(reportPath);
                        AttachFilesToLog(zipPath, testContext);

                    }
                }
            }
        }

        /// <summary>
        /// Method to initalise the class and define the report location - This should be called in [ClassInitialize] function
        /// </summary>
        public static void AssemblyInitialize()
        {
            //extentReport = new ExtentReports(reportPath, true);
            //extentReport.AddSystemInfo("Host Name", "Testhouse")
            //.AddSystemInfo("Environment", "QA")
            //.AddSystemInfo("User Name", "Testhouse");
            //extentReport.LoadConfig(projectPath + "\\Reports\\config.xml");
            //testr = projectPath;
        }

        /// <summary>
        /// Method to clean up the reports - This should be called in [ClassCleanup] function
        /// </summary>
        public static void AssemblyCleanup()
        {
            //extentReport.Flush();
            //extentReport.Close();
            //HideLogoVersion(reportPath);
        }

        public TestContext TestContext
        {
            get
            {
                return testContext;
            }

            set
            {
                testContext = value;
            }
        }
        private TestContext testContext;
    }
}
