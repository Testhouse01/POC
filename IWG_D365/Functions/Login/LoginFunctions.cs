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
using TesthouseSeleniumCSharp.Functions;
using TesthouseSeleniumCSharp.ObjectRepository;
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

namespace SeleniumCSharpMSTest.Functions
{
    public class LoginFunctions : Helper
    {
        Reader reader = new Reader();
        public String currentStatus;
        bool Flag = false;

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
        public void Login(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration, string hitUrl, string username, string password)
        {
            AddLog(driver, testInReport, testName, "Info", "Test started for " + username);
            driver.Navigate().GoToUrl(hitUrl);
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Application successfully launched", "ApplicationLaunched");
            driver.Manage().Window.Maximize();
            WaitUntil(driver, Control("emailAddress", "Generic"), 30);
            Element(driver, Control("emailAddress", "Generic")).SendKeys(username);
            WaitUntil(driver, Control("passWord", "Generic"), 60);
            Element(driver, Control("passWord", "Generic")).SendKeys(password);
            WaitUntil(driver, Control("signIn", "Generic"), 60);
            Element(driver, Control("signIn", "Generic")).Click();
            ThinkTime(3);
            HandlingEmailWarning(driver, testInReport, testName, testDataIteration);



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

            WaitUntil(driver, Control("emailAddress", "Generic"), 30);
            Element(driver, Control("emailAddress", "Generic")).SendKeys(username);
            WaitUntil(driver, Control("passWord", "Generic"), 60);
            Element(driver, Control("passWord", "Generic")).SendKeys(password);
            WaitUntil(driver, Control("signIn", "Generic"), 60);
            Element(driver, Control("signIn", "Generic")).Click();
            ThinkTime(3);

            HandlingEmailWarning(driver, testInReport, testName, testDataIteration);
            HandlingEditPopUp(driver, testInReport, testName, testDataIteration);

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
            try
            {
                HidingPureCloud(driver,testInReport,testName,testDataIteration);
            }
            catch(Exception e)
            {
                AddLog(driver, testInReport, testName, testDataIteration, "INFO", "No pop up handle", "popuphandle");
            }
            driver.SwitchTo().DefaultContent();
            WaitUntil(driver, Control("userInfoLink", "Generic"), 180);
            Element(driver, Control("userInfoLink", "Generic")).Click();
            WaitUntil(driver, Control("signOut", "Generic"), 180);
            MoveToElement(driver, Control("signOut", "Generic"));
            ElementHighlight(driver, Control("signOut", "Generic"));
            ActionsClick(driver, Control("signOut", "Generic"));
            // JSClick(driver, Control("signOut", "Generic"));
            AddLog(driver, testInReport, testName, testDataIteration, "Pass", "User logged out of application", "UserLogout");

        }
     
        
}











