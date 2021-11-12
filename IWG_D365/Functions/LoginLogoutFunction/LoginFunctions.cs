using OpenQA.Selenium;
using RelevantCodes.ExtentReports;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Support.UI;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SeleniumCSharpMSTest.GeneralFunctions;

using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using OpenQA.Selenium.Interactions;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Net.Mail;
using System.Net;
using System.Reflection;
using System.Globalization;
using System.Linq.Expressions;

namespace SeleniumCSharpMSTest.Functions
{
    public class LoginFunctions : Helper
    {
        Reader reader = new Reader();
        public String currentStatus;
        bool Flag = false;
        GenericFunctions generic = new GenericFunctions();

        ///Devi
        /// <summary>
        /// Method to Login to CRM Application
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="testInReport"></param>
        /// <param name="hitUrl"></param>
        /// <param name="testName"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// 
        public int Elements(IWebDriver driver, By by)
        {

            int elements = driver.FindElements(by).Count;
            return elements;
        }
        public void Login(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration, string hitUrl, string username, string password)
        {
            //ClearBrowserCache(driver, testInReport, testName, testDataIteration);

            AddLog(driver, testInReport, testName, "Info", "Test started for " + username);
            driver.Navigate().GoToUrl(hitUrl);
            string result = hitUrl.Substring(11, 5);
            AddLog(driver, testInReport, testName, testDataIteration, "Info", result + " Application successfully launched", result + " ApplicationLaunched");

            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Application successfully launched", "ApplicationLaunched");
            driver.Manage().Window.Maximize();
            ThinkTime(3);
            WaitUntil(driver, Control("emailAddress", "CommonObj"), 30);
            MoveToElement(driver, Control("emailAddress", "CommonObj"));
            Element(driver, Control("emailAddress", "CommonObj")).Click();
            Element(driver, Control("emailAddress", "CommonObj")).SendKeys(username);
            ThinkTime(2);
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "User Name Entered", "User Name Entered " + username);

            Element(driver, Control("emailAddress", "CommonObj")).SendKeys(Keys.Enter);
            try
            {
                ThinkTime(2);
                WaitUntil(driver, Control("passWord", "CommonObj"), 30);
                Element(driver, Control("passWord", "CommonObj")).Click();
                Element(driver, Control("passWord", "CommonObj")).SendKeys(password);
                AddLog(driver, testInReport, testName, testDataIteration, "Info", "Password Entered", "Password Entered" + username);
                ThinkTime(2);
                Element(driver, Control("passWord", "CommonObj")).SendKeys(Keys.Enter);
                ThinkTime(2);
                LoginExcep(driver, testInReport, testName, testDataIteration);
                Element(driver, Control("OkButton", "CommonObj")).Click();
                
                

                //HandlingEmailWarning(driver, testInReport, testName, testDataIteration);
                //ThinkTime(2);
            }
            catch (Exception e)
            {
                for (int m = 0; m < 10; m++)
                {
                    driver.Navigate().Refresh();
                    ThinkTime(3);
                    WaitUntil(driver, Control("emailAddress", "CommonObj"), 30);
                    MoveToElement(driver, Control("emailAddress", "CommonObj"));
                    Element(driver, Control("emailAddress", "CommonObj")).Click();
                    Element(driver, Control("emailAddress", "CommonObj")).SendKeys(username);
                    ThinkTime(2);
                    AddLog(driver, testInReport, testName, testDataIteration, "Info", "User Name Entered", "User Name Entered " + username);
                    Element(driver, Control("emailAddress", "CommonObj")).SendKeys(Keys.Enter);
                    ThinkTime(2);
                    AddLog(driver, testInReport, testName, testDataIteration, "Info", "Password Entry", "Password Entry " + m);
                    if (Elements(driver, Control("passWord", "CommonObj")) > 0)
                    {
                        break;
                    }
                }
                WaitUntil(driver, Control("passWord", "CommonObj"), 30);
                Element(driver, Control("passWord", "CommonObj")).Click();
                Element(driver, Control("passWord", "CommonObj")).SendKeys(password);
                AddLog(driver, testInReport, testName, testDataIteration, "Info", "Password Entered", "Password Entered" + username);
                ThinkTime(2);
                Element(driver, Control("passWord", "CommonObj")).SendKeys(Keys.Enter);
                ThinkTime(2);
                LoginExcep(driver, testInReport, testName, testDataIteration);               




            }
        }
        public void LoginExcep(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration)
        {
            ThinkTime(5);
            if (Elements(driver, Control("DiffAccnt", "CommonObj")) > 0)
            {
                Element(driver, Control("Nextbtn", "CommonObj")).Click();
                ThinkTime(2);
                WaitUntil(driver, Control("CancelLink", "CommonObj"), 60);
                Element(driver, Control("CancelLink", "CommonObj")).Click();

            }

        }


            public void LoginforUrl(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration, string hitUrl, string username, string password)
        {
            //ClearBrowserCache(driver, testInReport, testName, testDataIteration);

            AddLog(driver, testInReport, testName, "Info", "Test started for " + username);
            driver.Navigate().GoToUrl(hitUrl);
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Application successfully launched", "ApplicationLaunched");
            driver.Manage().Window.Maximize();
            ThinkTime(3);
        }
        public void LoginforUrl1(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration, string hitUrl, string username, string password)
        {
            //ClearBrowserCache(driver, testInReport, testName, testDataIteration);

            AddLog(driver, testInReport, testName, "Info", "Test started for " + username);
            driver.Navigate().GoToUrl(hitUrl);
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Application successfully launched", "ApplicationLaunched");
            driver.Manage().Window.Maximize();
            ThinkTime(3);
            driver.Navigate().Refresh();
            ThinkTime(3);
            WaitUntil(driver, Control("LogoMic", "CommonObj"), 30);
            Element(driver, Control("LogoMic", "CommonObj")).SendKeys(Keys.Enter);
            
            ThinkTime(3);
            WaitUntil(driver, Control("passWord", "CommonObj"), 30);
            Element(driver, Control("passWord", "CommonObj")).Click();
            Element(driver, Control("passWord", "CommonObj")).SendKeys(password);
            ThinkTime(2);
            Element(driver, Control("passWord", "CommonObj")).SendKeys(Keys.Enter);

        }



        /// <summary>
        /// Method to Login to CRM Application
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="testInReport"></param>
        /// <param name="hitUrl"></param>
        /// <param name="testName"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        public void LoginVerify(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration, string username, string password, string user)
        {

            driver.SwitchTo().DefaultContent();
            WaitUntil(driver, Control("userInfoLink", "Generic"), 60);
            Element(driver, Control("userInfoLink", "Generic")).Click();
            WaitUntil(driver, Control("loginVerifyValue1", "Generic"), 60);
            string userLogged = Element(driver, Control("loginVerifyValue1", "Generic")).Text;
            if (userLogged == user)
            {
                ThinkTime(1);
                AddLog(driver, testInReport, testName, testDataIteration, "Pass", "Successfully logged into the application as " + userLogged + "as expected", "Successfully logged in");
                ThinkTime(1);
                Element(driver, Control("userInfoLink", "Generic")).Click();
            }
            else
            {
                AddLog(driver, testInReport, testName, testDataIteration, "Fail", "Logged in to the application as wrong user: " + userLogged + " Expected user: " + user, "Logged in as wrong user");
                Assert.IsTrue(false);
            }
        }


        public void Loginafterlogout(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration, string username, string password)
        {


            WaitUntil(driver, Control("UseAnotherAccount", "Generic"), 240);
            Element(driver, Control("UseAnotherAccount", "Generic")).Click();
            AddLog(driver, testInReport, testName, "Info", "Test started for " + username);

            WaitUntil(driver, Control("emailAddress", "CommonObj"), 30);
            MoveToElement(driver, Control("emailAddress", "CommonObj"));
            Element(driver, Control("emailAddress", "CommonObj")).Click();
            Element(driver, Control("emailAddress", "CommonObj")).SendKeys(username);
            ThinkTime(2);
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "User Name Entered", "User Name Entered " + username);

            Element(driver, Control("emailAddress", "CommonObj")).SendKeys(Keys.Enter);
            try
            {
                ThinkTime(2);
                WaitUntil(driver, Control("passWord", "CommonObj"), 30);
                Element(driver, Control("passWord", "CommonObj")).Click();
                Element(driver, Control("passWord", "CommonObj")).SendKeys(password);
                AddLog(driver, testInReport, testName, testDataIteration, "Info", "Password Entered", "Password Entered" + username);
                ThinkTime(2);
                Element(driver, Control("passWord", "CommonObj")).SendKeys(Keys.Enter);
                ThinkTime(2);
                //HandlingEmailWarning(driver, testInReport, testName, testDataIteration);
                //ThinkTime(2);
                ThinkTime(2);
                LoginExcep(driver, testInReport, testName, testDataIteration);
            }
            catch (Exception e)
            {
                for (int m = 0; m < 10; m++)
                {
                    driver.Navigate().Refresh();
                    ThinkTime(3);
                    WaitUntil(driver, Control("emailAddress", "CommonObj"), 30);
                    MoveToElement(driver, Control("emailAddress", "CommonObj"));
                    Element(driver, Control("emailAddress", "CommonObj")).Click();
                    Element(driver, Control("emailAddress", "CommonObj")).SendKeys(username);
                    ThinkTime(2);
                    AddLog(driver, testInReport, testName, testDataIteration, "Info", "User Name Entered", "User Name Entered " + username);
                    Element(driver, Control("emailAddress", "CommonObj")).SendKeys(Keys.Enter);
                    ThinkTime(2);
                    AddLog(driver, testInReport, testName, testDataIteration, "Info", "Password Entry", "Password Entry " + m);
                    if (Elements(driver, Control("passWord", "CommonObj")) > 0)
                    {
                        break;
                    }
                }
                WaitUntil(driver, Control("passWord", "CommonObj"), 30);
                Element(driver, Control("passWord", "CommonObj")).Click();
                Element(driver, Control("passWord", "CommonObj")).SendKeys(password);
                AddLog(driver, testInReport, testName, testDataIteration, "Info", "Password Entered", "Password Entered" + username);
                ThinkTime(2);
                Element(driver, Control("passWord", "CommonObj")).SendKeys(Keys.Enter);
                ThinkTime(2);
                //HandlingEmailWarning(driver, testInReport, testName, testDataIteration);    
                ThinkTime(2);
                LoginExcep(driver, testInReport, testName, testDataIteration);
            }
        }



            /// <summary>
            /// Method to Logout of CRM Application
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="testInReport"></param>
            /// <param name="testName"></param>
            /// <param name="testDataIteration"></param>
            public void Logout(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration)
        {

           // ThinkTime(5);
            //try
            //{
            //    generic.HidingPureCloud(driver, testInReport, testName, testDataIteration);
            //}
            //catch (Exception e)
            //{
            //    AddLog(driver, testInReport, testName, testDataIteration, "INFO", "No pop up handle", "popuphandle");
            //}
            ThinkTime(10);
            driver.SwitchTo().DefaultContent();
            WaitUntil(driver, Control("userInfoLink", "CommonObj"), 180);
            Element(driver, Control("userInfoLink", "CommonObj")).Click();
            ThinkTime(2);
            WaitUntil(driver, Control("signOut", "CommonObj"), 180);
            MoveToElement(driver, Control("signOut", "CommonObj"));
            ElementHighlight(driver, Control("signOut", "CommonObj"));
            ActionsClick(driver, Control("signOut", "CommonObj"));
            // JSClick(driver, Control("signOut", "Generic"));
            AddLog(driver, testInReport, testName, testDataIteration, "Pass", "User logged out of application", "UserLogout");

        }





        /// <summary>
        /// Method to Logout from SharePoint Application
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="testInReport"></param>
        /// <param name="testName"></param>
        /// <param name="testDataIteration"></param>
        public void logoutSharePoint(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration)
        {
            driver.SwitchTo().DefaultContent();
            ThinkTime(5);
            WaitUntil(driver, Control("o365UserInfo", "sharePoint"), 180);
            Element(driver, Control("o365UserInfo", "sharePoint")).Click();
            WaitUntil(driver, Control("o365SignOut", "sharePoint"), 180);
            MoveToElement(driver, Control("o365SignOut", "sharePoint"));
            ElementHighlight(driver, Control("o365SignOut", "sharePoint"));
            ActionsClick(driver, Control("o365SignOut", "sharePoint"));
            AddLog(driver, testInReport, testName, testDataIteration, "Pass", "User logged out of SharePoint application", "UserLogout");
            //driver.Quit();
        }

    }
    }











