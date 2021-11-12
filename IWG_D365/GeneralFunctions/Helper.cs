using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Threading;
using RelevantCodes.ExtentReports;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Net;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;
using System.Collections.Generic;

namespace SeleniumCSharpMSTest.GeneralFunctions
{
    /// <summary>
    /// Helper class for general functions that can be used across the project and helping methods
    /// </summary>
    public class Helper
    {
        private static string runResults;
        private static string actualPath = "";
        private static string resultDate;
        Reader reader = new Reader();
        public static List<string> screenshotsForVSTSUpload = new List<string>();
        public static string zipPath = "";
        /// <summary>
        /// Method to wait
        /// </summary>
        /// <param name="wait"></param>
        public static void ThinkTime(int wait)
        {
            Thread.Sleep(wait * 1000);
        }

        /// <summary>
        /// Simple wait and wait up to timeout wait
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="by"></param>
        /// <param name="wait"></param>
        /// <param name="timeOutWait"></param>
        /// <returns></returns>
        public static int TimeoutWait(IWebDriver driver, By by, int wait, int timeOutWait)
        {
            int flag = 0;
            try
            {
                Thread.Sleep(wait);
                Element(driver, by).Click();
            }
            catch (Exception)
            {
                try
                {
                    Thread.Sleep(timeOutWait);
                    Element(driver, by).Click();
                }
                catch (Exception f)
                {
                    flag = 1;
                    throw f;
                }
            }
            return flag;
        }

        /// <summary>
        /// Wait before frame and check availability of element in frame
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="frameIndex"></param>
        /// <param name="by"></param>
        /// <param name="wait"></param>
        /// <param name="timeOutWait"></param>
        public static void WaitBeforeFrameAndCheck(IWebDriver driver, int frameIndex, By by, int wait, int timeOutWait)
        {
            try
            {
                Thread.Sleep(wait);
                driver.SwitchTo().Frame(frameIndex);
                WaitUntil(driver, by, 1);
                driver.SwitchTo().DefaultContent();
            }
            catch (NoSuchElementException)
            {
                try
                {
                    Thread.Sleep(timeOutWait);
                    driver.SwitchTo().Frame(frameIndex);
                    WaitUntil(driver, by, 1);
                    driver.SwitchTo().DefaultContent();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Timeout: " + e);
                }
            }
        }

        /// <summary>
        /// Click and wait
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="by"></param>
        /// <param name="timeOut"></param>
        public static void WaitAndClick(IWebDriver driver, By by, int timeOut)
        {
            while (timeOut > 0)
            {
                try
                {
                    if (Element(driver, by).Displayed)
                    {
                        Element(driver, by).Click();
                    }
                }
                catch (StaleElementReferenceException)
                {
                    Thread.Sleep(1000);
                }
                catch (InvalidOperationException)
                {
                    Thread.Sleep(1000);

                }
                catch (NoSuchElementException)
                {
                    try
                    {
                        if (!Element(driver, by).Displayed)
                        {

                        }
                    }
                    catch (Exception)
                    {
                    }
                }
                timeOut--;
            }
        }

        /// <summary>
        /// To take screenshot of the browser page
        /// </summary>
        /// <param name="driver"></param>
        /// <returns></returns>
        public static string GetScreenShotUnexpected(IWebDriver driver)
        {
            ITakesScreenshot takeScreenshot = (ITakesScreenshot)driver;
            Screenshot screenshot = takeScreenshot.GetScreenshot();
            DateTime daTime = DateTime.Now;
            string startupPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase))));
            string resultDate = daTime.ToString("dd-MMM-yy_HH.mm.ss");
            string finalPath = startupPath + @"\Reports\ScreenShots\" + "screenShotName" + resultDate + ".png";
            string localPath = new Uri(finalPath).LocalPath;
            screenshot.SaveAsFile(localPath, ScreenshotImageFormat.Png);
            return localPath;
        }

        /// <summary>
        /// To set the Project path
        /// </summary>
        /// <returns></returns>
        public static string SetProjectPath()
        {
            string path = Assembly.GetCallingAssembly().CodeBase;
            string actualPathTemp = "";
            if (path.Contains("bin"))
            {
                actualPathTemp = path.Substring(0, path.LastIndexOf("bin"));
            }
            else if (path.Contains("TestResults"))
            {
                string executingAssemblyName = Path.GetFileName(Assembly.GetExecutingAssembly().Location);
                actualPathTemp = path.Substring(0, path.LastIndexOf("TestResults")) + executingAssemblyName.Replace(".dll", "");
            }

            if (actualPathTemp.Contains("TesthouseSeleniumCSharp"))
            {
                actualPath = actualPathTemp.Replace("TesthouseSeleniumCSharp", ConfigurationManager.AppSettings["Project"]);
            }
            else
            {
                actualPath = actualPathTemp;
            }

            string projectPath = new Uri(actualPath).LocalPath;
            return projectPath;

        }

        /// <summary>
        /// To set the report path using project path
        /// </summary>
        /// <param name="projectPath"></param>
        /// <returns></returns>

        public static string SetReport()
        {
            Helper.resultDate = DateTime.Now.ToString("dd-MMM-yy_HH.mm.ss");
            string localPath = new Uri(Helper.SetProjectPath()).LocalPath;
            Directory.CreateDirectory(localPath + "\\Reports\\" + Helper.resultDate + "_run results");
            string str = localPath + "\\Reports\\" + Helper.resultDate + "_run results\\Results.html";
            Helper.runResults = str;
            return str;
        }

        /// <summary>
        /// To take screenshot of the browser with screenshot name stored in testname folder
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="screenshotName"></param>
        /// <param name="testName"></param>
        /// <param name="testDataIteration"></param>
        /// <returns></returns>
       public static string GetScreenshot(
      IWebDriver driver,
      string screenshotName,
      string testName,
      string testDataIteration)
    {
      Screenshot screenshot = ((ITakesScreenshot) driver).GetScreenshot();
      string path = Path.GetDirectoryName(Helper.runResults) + "/" + testName + "/Iteration" + testDataIteration + "\\";
      Directory.CreateDirectory(path);
      string localPath = new Uri(path + screenshotName + ".png").LocalPath;
      screenshot.SaveAsFile(localPath, (ScreenshotImageFormat) 0);
      return "..\\" + Helper.resultDate + "_run results\\" + testName + "\\Iteration" + testDataIteration + "\\" + screenshotName + ".png";
    }

        /// <summary>
        /// To get current method name
        /// </summary>
        /// <returns></returns>
        public static string GetCurrentMethod()
        {
            StackTrace stackTrace = new StackTrace();
            StackFrame stackFrame = stackTrace.GetFrame(1);

            return stackFrame.GetMethod().Name;
        }

        /// <summary>
        /// Return web element using By control properties
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="object1"></param>
        /// <returns></returns>
        public static IWebElement Element(IWebDriver driver, By by)
        {
            IWebElement element = driver.FindElement(by);
            return element;
        }

        /// <summary>
        /// Wait until the element is found
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="by"></param>
        /// <param name="timeoutInSeconds"></param>
        /// <returns></returns>
        public static IWebElement WaitUntil(IWebDriver driver, By by, int timeOutInSeconds)
        {
            if (timeOutInSeconds > 0)
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeOutInSeconds));
                return wait.Until(drv =>drv.FindElement(by));
            }
            return driver.FindElement(by);
        }

        /// <summary>
        /// Actions click
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="elementToClick"></param>
        public static void ActionsClick(IWebDriver driver, By by)
        {
            Actions action = new Actions(driver);
            action.Click(Element(driver, by)).Build().Perform();
        }

        /// <summary>
        /// Wait until the Element is visible
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="locator"></param>
        /// <param name="maxSeconds"></param>
        /// <returns></returns>
        public static IWebElement WaitUntilIElementVisible(IWebDriver driver, By locator, int maxSeconds)
        {
            return new WebDriverWait(driver, TimeSpan.FromSeconds(maxSeconds)).Until(ExpectedConditions.ElementExists(locator));
        }

        /// <summary>
        /// Select elements from drop down
        /// </summary>
        /// <param name="findElement"></param>
        /// <returns></returns>
        public static SelectElement Select(IWebElement findElement)
        {
            SelectElement select = new SelectElement(findElement);
            return select;
        }

        /// <summary>
        /// Switch to Frames
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="frame"></param>
        public static void SwitchToFrame(IWebDriver driver, string frame)
        {
            driver.SwitchTo().Frame(frame);
        }

        /// <summary>
        /// Switch to Frames using different identification methods
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="by"></param>
        public static void SwitchToFrameElement(IWebDriver driver, By by)
        {
            driver.SwitchTo().Frame(driver.FindElement(by));
        }

        /// <summary>
        /// Highlight Elements in Web page
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="by"></param>
        public void ElementHighlight(IWebDriver driver, By by)
        {
            for (int i = 0; i < 2; i++)
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                js.ExecuteScript(
                        "arguments[0].setAttribute('style', arguments[1]);",
                        Element(driver, by), "color: red; border: 3px solid red;");
                js.ExecuteScript(
                        "arguments[0].setAttribute('style', arguments[1]);",
                        Element(driver, by), "");
            }
        }

        /// <summary>
        /// Method to wait for element until it is enable
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="by"></param>
        public void WaitForElementEnabled(IWebDriver driver, By by)
        {
            for (int i = 1; i <= 100; i++)
            {
                if (!Element(driver, by).Enabled)
                {
                    ThinkTime(2);
                }
                else
                {
                    break;
                }
            }
        }

        /// <summary>
        /// Method to wait for element until the Element is not present in DOM
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="by"></param>
        public void WaitForElementExists(IWebDriver driver, By by)
        {
            for (int i = 1; i <= 100; i++)
            {
                if (driver.FindElements(by).Count != 0)
                {
                    ThinkTime(2);
                }
                else
                {
                    break;
                }
            }
        }

        /// <summary>
        /// Method to wait for element until the Element is not present in DOM
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="by"></param>
        public void WaitForElementNotExists(IWebDriver driver, By by)
        {
            for (int i = 1; i <= 100; i++)

            {
                if (driver.FindElements(by).Count == 0)
                {
                    break;
                }
                else
                {
                    ThinkTime(2);
                }
            }
        }

        /// <summary>
        /// Send Keys to element using Actions
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="by"></param>
        /// <param name="data"></param>
        public void ActionSendKeys(IWebDriver driver, By by, string textString)
        {
            Actions actions = new Actions(driver);
            IWebElement control = Element(driver, by);
            actions.MoveToElement(control);
            actions.Click();
            actions.SendKeys(textString);
            actions.Build().Perform();
        }

        /// <summary>
        /// Returns the By property for the object values stored in Excel.
        /// </summary>
        /// <param name="controlName"> Control Name in the Object Repository Sheet</param>
        /// <param name="objectSheet">Excel sheet where the control is defined</param>
        /// <returns></returns>

        public By Control(string controlName, string objectSheet)
        {
            By control = null;
            Keywords KeyFound = reader.FindControlinList(controlName, objectSheet);
            string propertyName = KeyFound.PropertyName;
            string propertyValue = KeyFound.PropertyValue;

            switch (propertyName)
            {
                case "Id":
                    control = By.Id(propertyValue);
                    break;
                case "Css":
                    control = By.CssSelector(propertyValue);
                    break;
                case "TagName":
                    control = By.TagName(propertyValue);
                    break;
                case "Name":
                    control = By.Name(propertyValue);
                    break;
                case "ClassName":
                    control = By.ClassName(propertyValue);
                    break;
                case "LinkText":
                    control = By.LinkText(propertyValue);
                    break;
                case "PartialLinkText":
                    control = By.PartialLinkText(propertyValue);
                    break;
                case "XPath":
                    control = By.XPath(propertyValue);
                    break;
            }
            return control;
        }

        /// <summary>
        /// Returns the By property for the object values stored in Excel.
        /// </summary>
        /// <param name="controlName">Control Name in the Object Repository Sheet</param>
        /// <param name="propertyValueExcel">Test Data from test Data excel file</param>
        /// <param name="objectSheet">Excel sheet where the control is defined</param>
        /// <returns></returns>
        public By Control(string controlName, string propertyValueExcel, string objectSheet)
        {
            By control = null;
            Keywords KeyFound = reader.FindControlinList(controlName, objectSheet);
            string propertyValuenew = "";
            string propertyName = KeyFound.PropertyName;
            string propertyValue = KeyFound.PropertyValue;
            propertyValuenew = propertyValue.Replace("+Data+", propertyValueExcel);

            switch (propertyName)
            {
                case "Name":
                    control = By.Name(propertyValueExcel);
                    break;
                case "ClassName":
                    control = By.ClassName(propertyValueExcel);
                    break;
                case "LinkText":
                    control = By.LinkText(propertyValueExcel);
                    break;
                case "PartialLinkText":
                    control = By.PartialLinkText(propertyValueExcel);
                    break;
                case "XPath":
                    control = By.XPath(propertyValuenew);
                    break;
            }
            return control;
        }


        /// <summary>
        /// Returns the By property for the object values stored in Excel.
        /// </summary>
        /// <param name="controlName">Control Name in the Object Repository Sheet</param>
        /// <param name="propertyValueExcel1"></param>
        /// <param name="propertyValueExcel2"></param>
        /// <param name="objectSheet"></param>
        /// <returns></returns>

        public By Control(string controlName, string propertyValueExcel1, string propertyValueExcel2, string objectSheet)
        {
            By control = null;
            Keywords KeyFound = reader.FindControlinList(controlName, objectSheet);
            string propertyValuenew1 = "";
            string propertyValuenew2 = "";
            string propertyName = KeyFound.PropertyName;
            string propertyValue = KeyFound.PropertyValue;
            propertyValuenew1 = propertyValue.Replace("+Data+", propertyValueExcel1);
            propertyValuenew2 = propertyValuenew1.Replace("+Data1+", propertyValueExcel2);
            switch (propertyName)
            {
                case "XPath":
                    control = By.XPath(propertyValuenew2);
                    break;
            }
            return control;
        }

        /// <summary>
        /// Run Java script code for controls stored in excel file
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="controlName">Control Name in the Object Repository Sheet</param>
        /// <param name="propertyValueExcel">Test Data from test Data excel file</param>
        /// <param name="objectSheet">Excel sheet where the control is defined</param>
        public void RunJavaScript(IWebDriver driver, string controlName, string propertyValueExcel, string objectSheet)
        {
            Keywords KeyFound = reader.FindControlinList(controlName, objectSheet);
            string propertyValueNew = "";
            string propertyName = KeyFound.PropertyName;
            string propertyValue = KeyFound.PropertyValue;
            propertyValueNew = propertyValue.Replace("+Data+", propertyValueExcel);

            switch (propertyName)
            {
                case "JavaScript":
                    ((IJavaScriptExecutor)driver).ExecuteScript(propertyValueNew);
                    break;
            }
        }
/// <summary>
/// ////////////////////////////Excel reader ////////////////////
/// </summary>
/// <returns></returns>
        public static string setReportPath()
        {
            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actualPath = path.Substring(0, path.LastIndexOf("bin"));
            string projectPath = new Uri(actualPath).LocalPath;
            return projectPath;

        }

        public static string dataFilePath(string file)
        {
            ExcelToTable env = new ExcelToTable();
            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actualPath = path.Substring(0, path.LastIndexOf("bin"));
            string projectPath = new Uri(actualPath).LocalPath;
            string filePath = projectPath + file;
            return filePath;
        }


        public static List<DataRow> getTestCaseList(string sheetName)
        {
            //ExcelToTable d = new ExcelToTable();
            // d.PopulateInCollection(dataFilePath(), ConfigurationManager.AppSettings["TestSheetName"], thisTestName);
          ExcelDataAccess.setExcelInfo(sheetName, dataFilePath(ConfigurationManager.AppSettings["TestData"]), setReportPath());
            return ExcelDataAccess.getTestCaseData(sheetName, dataFilePath(ConfigurationManager.AppSettings["TestData"]));

        }

        ///////////////////////////


        /// <summary>
        /// This function helps to log steps in report with screenshot of the browser. It saves the captured screenshot with a unique filename (by appending the time stamp)<para />
        /// *Result* - (Not case sensitive) "Pass" or "Fail" or "Info" or "Error" or "Fatal" or "Skip" or "Unknown" or "Warning" <para />
        /// *stepName* -Step Description and Tag-line of the screenshot in report<para />
        /// *fileName* -File name of the screenshot ( do not use any special characters here other than underscore(_) )<para />
        /// </summary>
        /// <param name="result"> Result can be Pass/Fail/Info/Error/Fatal/Skip/Unknown/Warning</param>
        /// <param name="stepName">Step Description and Tag-line of the screenshot</param>
        /// <param name="fileName">File name of the screenshot</param>
        /// <param name="driver">Object for webdriver instance</param>
        /// <param name="testInReport">Test in the report</param>
        /// <param name="testName">Method name from Testing.cs</param>
        /// <param name="testDataIteration"></param>

        public void AddLog(
          IWebDriver driver,
          ExtentTest testInReport,
          string testName,
          string testDataIteration,
          string result,
          string stepName,
          string fileName)
        {
            Helper.ThinkTime(1);
            string str = DateTime.Now.ToString("HH.mm.ss");
            string screenshotName = fileName + "_" + str;
            if (string.Equals(result, "Pass", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)0, stepName, testInReport.AddScreenCapture(Helper.GetScreenshot(driver, screenshotName, testName, testDataIteration)) + "Filename : " + screenshotName + ".png");
            else if (string.Equals(result, "Fail", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)1, stepName, testInReport.AddScreenCapture(Helper.GetScreenshot(driver, screenshotName, testName, testDataIteration)) + "Filename : " + screenshotName + ".png");
            else if (string.Equals(result, "Info", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)5, stepName, testInReport.AddScreenCapture(Helper.GetScreenshot(driver, screenshotName, testName, testDataIteration)) + "Filename : " + screenshotName + ".png");
            else if (string.Equals(result, "Error", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)3, stepName, testInReport.AddScreenCapture(Helper.GetScreenshot(driver, screenshotName, testName, testDataIteration)) + "Filename : " + screenshotName + ".png");
            else if (string.Equals(result, "Fatal", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)2, stepName, testInReport.AddScreenCapture(Helper.GetScreenshot(driver, screenshotName, testName, testDataIteration)) + "Filename : " + screenshotName + ".png");
            else if (string.Equals(result, "Skip", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)6, stepName, testInReport.AddScreenCapture(Helper.GetScreenshot(driver, screenshotName, testName, testDataIteration)) + "Filename : " + screenshotName + ".png");
            else if (string.Equals(result, "Unknown", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)7, stepName, testInReport.AddScreenCapture(Helper.GetScreenshot(driver, screenshotName, testName, testDataIteration)) + "Filename : " + screenshotName + ".png");
            else if (string.Equals(result, "Warning", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)4, stepName, testInReport.AddScreenCapture(Helper.GetScreenshot(driver, screenshotName, testName, testDataIteration)) + "Filename : " + screenshotName + ".png");
            this.StoreScreenshotsForVSTSUpload(Helper.SetProjectPath(), Helper.GetScreenshot(driver, screenshotName, testName, testDataIteration));
        }

        public static void AddLog(
          IWebDriver driver,
          ExtentTest testInReport,
          string testName,
          string result,
          string stepName)
        {
            Helper.ThinkTime(1);
            if (string.Equals(result, "Pass", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)0, stepName, "");
            else if (string.Equals(result, "Fail", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)1, stepName, "");
            else if (string.Equals(result, "Info", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)5, stepName, "");
            else if (string.Equals(result, "Error", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)3, stepName, "");
            else if (string.Equals(result, "Fatal", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)2, stepName, "");
            else if (string.Equals(result, "Skip", StringComparison.OrdinalIgnoreCase))
                testInReport.Log((LogStatus)6, stepName, "");
            else if (string.Equals(result, "Unknown", StringComparison.OrdinalIgnoreCase))
            {
                testInReport.Log((LogStatus)7, stepName, "");
            }
            else
            {
                if (!string.Equals(result, "Warning", StringComparison.OrdinalIgnoreCase))
                    return;
                testInReport.Log((LogStatus)4, stepName, "");
            }
        }

        public static void AddUpdateAppSettings(string key, string value)
        {
            try
            {
                System.Configuration.Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                KeyValueConfigurationCollection settings = configuration.AppSettings.Settings;
                if (settings[key] == null)
                    settings.Add(key, value);
                else
                    settings[key].Value = value;
                configuration.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configuration.AppSettings.SectionInformation.Name);
            }
            catch (ConfigurationErrorsException ex)
            {
                Console.WriteLine("Error writing app settings");
            }
        }

        public static void SendEmail(string mailPath)
        {
            try
            {
                MailMessage message = new MailMessage();
                SmtpClient smtpClient = new SmtpClient(ConfigurationManager.AppSettings["SMTPServer"]);
                smtpClient.Timeout = 150000;
                message.From = new MailAddress(ConfigurationManager.AppSettings["SenderEmail"]);
                message.To.Add(ConfigurationManager.AppSettings["ReceiverEmail"]);
                if (ConfigurationManager.AppSettings["BccEmail"] != "")
                    message.Bcc.Add(ConfigurationManager.AppSettings["BccEmail"]);
                if (ConfigurationManager.AppSettings["CCEmail"] != "")
                    message.CC.Add(ConfigurationManager.AppSettings["CCEmail"]);
                string str1 = DateTime.Now.ToString("dd-MMM-yyyy");
                message.Subject = "Automation Test Execution Report : " + str1;
                message.Body = "Dear All, \n\nPlease find the attached test execution report for date : " + str1 + " \n\nThe location of the original report is in the below location : " + mailPath + "\n\nNote: The report does not contain screenshots.To view the reports with screenshots, kindly open the report from the location specified above. \n\nRegards, \nTesthouse Automation Team";
                string str2 = mailPath.Replace("Results.html", "") + "\\EmailResults.html";
                System.IO.File.Copy(mailPath, str2);
                Helper.RemoveStepDetails(str2);
                Helper.DashboardToMail(str2);
                Attachment attachment = new Attachment(str2);
                message.Attachments.Add(attachment);
                smtpClient.Port = int.Parse(ConfigurationManager.AppSettings["SMTPPort"]);
                smtpClient.Credentials = (ICredentialsByHost)new NetworkCredential(ConfigurationManager.AppSettings["SenderEmail"], ConfigurationManager.AppSettings["SenderEmailPassword"]);
                smtpClient.EnableSsl = true;
                smtpClient.Send(message);
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static void HideLogoVersion(string mailPath)
        {
            string contents1 = new Regex("<span>ExtentReports</span>").Replace(System.IO.File.ReadAllText(mailPath), "<span>TesthouseReports</span>");
            System.IO.File.WriteAllText(mailPath, contents1);
            string contents2 = new Regex("<div class='logo-container'>").Replace(System.IO.File.ReadAllText(mailPath), "<div class='logo-container' style='display: none;'>");
            System.IO.File.WriteAllText(mailPath, contents2);
            string contents3 = new Regex("<span>v(.*?)</span>").Replace(System.IO.File.ReadAllText(mailPath), "");
            System.IO.File.WriteAllText(mailPath, contents3);
        }

        public static void RemoveStepDetails(string emailReportPath)
        {
            string contents = new Regex("<td class='step-details'>(.*?)</td>").Replace(System.IO.File.ReadAllText(emailReportPath), "");
            System.IO.File.WriteAllText(emailReportPath, contents);
        }

        public static void DashboardToMail(string emailReportPath)
        {
            string contents1 = new Regex("<div id='test-view' class='row _addedTable'>").Replace(System.IO.File.ReadAllText(emailReportPath), "<div id='test-view' class='row _addedTable hide'>");
            System.IO.File.WriteAllText(emailReportPath, contents1);
            string contents2 = new Regex("<li class='analysis waves-effect'><a href='#!' class='categories-view'").Replace(System.IO.File.ReadAllText(emailReportPath), "<li class='analysis waves-effect' style='display: none;'><a href='#!' class='categories-view'");
            System.IO.File.WriteAllText(emailReportPath, contents2);
            Regex regex1 = new Regex("<li class='analysis waves-effect active'><a href='#!' class='test-view'");
            string input1 = System.IO.File.ReadAllText(emailReportPath);
            string contents3 = regex1.Replace(input1, "<li class='analysis waves-effect' style='display: none;'><a href='#!' class='test-view'");
            System.IO.File.WriteAllText(emailReportPath, contents3);
            string contents4 = new Regex("<li class='analysis waves-effect'><a href='#!' class='dashboard-view'").Replace(System.IO.File.ReadAllText(emailReportPath), "<li class='analysis waves-effect active'><a href='#!' class='dashboard-view'");
            System.IO.File.WriteAllText(emailReportPath, contents4);
            Regex regex2 = new Regex("<li class='analysis waves-effect'><a href='#!' class='exceptions-view'");
            string input2 = System.IO.File.ReadAllText(emailReportPath);
            string contents5 = regex1.Replace(input2, "<li class='analysis waves-effect' style='display: none;'><a href='#!' class='exceptions-view'");
            System.IO.File.WriteAllText(emailReportPath, contents5);
            string contents6 = new Regex("<body class='extent standard hide-overflow'>").Replace(System.IO.File.ReadAllText(emailReportPath), "<body class='extent standard hide-overflow'><script> window.onload = function() {   $('.dashboard-view').trigger('click'); };</script>");
            System.IO.File.WriteAllText(emailReportPath, contents6);
        }

        public void ReRun(
          string testName,
          ref int maxTestRuns,
          int maxTestRunsCount,
          ExtentTest extentTest,
          RelevantCodes.ExtentReports.ExtentReports extentReport,
          Type classType,
          object objTestFail,
          string testInitializeName = null,
          string testCleanupName = null)
        {
            if (maxTestRuns > 0 && maxTestRuns <= maxTestRunsCount)
            {
                --maxTestRuns;
                if (testCleanupName != null)
                    classType.GetMethod(testCleanupName).Invoke(objTestFail, (object[])null);
                else
                    extentReport.EndTest(extentTest);
                extentTest.Log((LogStatus)5, string.Format("Running Again, Retry count:{0} ", (object)(maxTestRunsCount - maxTestRuns)));
                if (testInitializeName != null)
                    classType.GetMethod(testInitializeName).Invoke(objTestFail, (object[])null);
                classType.GetMethod(testName).Invoke(objTestFail, (object[])null);
            }
            else
            {
                extentTest.Log((LogStatus)1, string.Format("Failed even after retrying {0} times", (object)maxTestRunsCount));
                Assert.Fail("Failed even after retrying");
            }
        }

        public void ActionsDoubleClick(IWebDriver driver, By by)
        {
            Actions actions = new Actions(driver);
            IWebElement iwebElement = Helper.Element(driver, by);
            actions.MoveToElement(iwebElement);
            actions.DoubleClick();
            actions.Build().Perform();
        }

        public void JSClick(IWebDriver driver, By by)
        {
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", new object[1]
            {
        (object) Helper.Element(driver, by)
            });
        }

        public void KillChromeInstances()
        {
            try
            {
                foreach (Process process in Process.GetProcessesByName("chromedriver"))
                    process.Kill();
            }
            catch (Exception ex)
            {
            }
        }

        public static void KillExcel()
        {
            foreach (Process process in Process.GetProcessesByName("EXCEL.EXE *32"))
                process.Kill();
        }

        public void MoveToElement(IWebDriver driver, By by)
        {
            Actions actions = new Actions(driver);
            IWebElement iwebElement = Helper.Element(driver, by);
            actions.MoveToElement(iwebElement);
            actions.Build().Perform();
        }

        /// <summary>
        /// Method to wait for alert to be present
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="timeOutInSeconds"></param>
        public void WaitForAlert(IWebDriver driver, int timeOutInSeconds)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeOutInSeconds));
            wait.Until(ExpectedConditions.AlertIsPresent());
        }

        public void UploadScreenshots(TestContext testContextInstanceForVSTS)
        {
            try
            {
                foreach (string pathOfTheFileToUpload in Helper.screenshotsForVSTSUpload)
                    this.AttachFilesToLog(pathOfTheFileToUpload, testContextInstanceForVSTS);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while adding screenshots to log");
                Console.WriteLine(ex.Message);
            }
        }

        public void AttachFilesToLog(
          string pathOfTheFileToUpload,
          TestContext testContextInstanceForVSTS)
        {
            try
            {
                testContextInstanceForVSTS.AddResultFile(pathOfTheFileToUpload);
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("pathOfTheFileToUpload - {0}", (object)pathOfTheFileToUpload));
                Console.WriteLine("Error while adding file to log file");
                Console.WriteLine(ex.Message);
            }
        }

        public void AddReportFileToZip(string reportPathFileName)
        {
            try
            {
                string str = reportPathFileName.Substring(0, reportPathFileName.LastIndexOf("Results.html"));
                string fileName = Path.GetFileName(Directory.GetParent(str).FullName);
                Helper.zipPath = str.Substring(0, str.LastIndexOf(Path.GetFileName(Directory.GetParent(str).FullName))) + string.Format("{0}_zip.zip", (object)fileName);
                if (System.IO.File.Exists(Helper.zipPath))
                    System.IO.File.Delete(Helper.zipPath);
                ZipFile.CreateFromDirectory(str, Helper.zipPath, CompressionLevel.Fastest, true);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while creating zip folder from Results.html file path");
                Console.WriteLine(ex.Message);
            }
        }

        public void StoreScreenshotsForVSTSUpload(string pathTillReports, string screenshotPath)
        {
            string str = pathTillReports + "\\Reports\\" + screenshotPath.Substring(3);
            Helper.screenshotsForVSTSUpload.Add(str);
        }

    }
}

