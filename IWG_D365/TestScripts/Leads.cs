using Microsoft.VisualStudio.TestTools.UnitTesting;
using SeleniumCSharpMSTest.Functions;
using System;
using System.Configuration;
using System.Linq;
using SeleniumCSharpMSTest.GeneralFunctions;
using OpenQA.Selenium;
using System.Globalization;
using System.Collections.Generic;
using System.Data;

namespace SeleniumCSharpMSTest.TestScripts
{

    [TestClass]
    
    public class Leads:BaseClass
    {

        LoginFunctions login = new LoginFunctions();
        GenericFunctions generic = new GenericFunctions();


        /// <summary>
        /// Verify the quick create contact details are same as lead details.
        /// </summary>
        [TestCategory("smoke"), TestCategory("InProgress"), TestCategory("NoRework")]
        [TestProperty("TestcaseID", "RTA-5260")]
        public static IEnumerable<object[]> RTA_5848_Data()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
[DataTestMethod]
[DynamicData(nameof(RTA_5848_Data), DynamicDataSourceType.Method)]
        public void SalesAgentLeadVerification(DataRow Ro)
        {
            try
            {
                string now = System.DateTime.Now.ToString();
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.SelectNewTab(driver, extentTest, testName, testDataIteration);
                generic.HidingPureCloud(driver, extentTest, testName, testDataIteration);
                generic.SelectAndVerifyLeadsFieldsSalesAgent(driver, extentTest, testName, testDataIteration, Ro["Lastname"].ToString(), Ro["Email"].ToString(), Ro["BussinessPhone"].ToString(), Ro["Source"].ToString(), Ro["MajorSource"].ToString(), Ro["MinorSource"].ToString(),now);
                generic.scrollUpPage(driver, extentTest, testName, testDataIteration);
                generic.AddNewContact(driver, extentTest, testName, testDataIteration);
                generic.VerifyContactField(driver, extentTest, testName, testDataIteration, Ro["Lastname"].ToString(), Ro["Email"].ToString(), Ro["BussinessPhone"].ToString(), Ro["Language"].ToString(),now);
                generic.SelectAndVerifyLeadsFieldsSalesAgent(driver, extentTest, testName, testDataIteration, Ro["Lastname1"].ToString(), Ro["Email1"].ToString(), Ro["BussinessPhone"].ToString(), Ro["Source"].ToString(), Ro["MajorSource"].ToString(), Ro["MinorSource"].ToString(),now);
                generic.scrollUpPage(driver, extentTest, testName, testDataIteration);
                generic.AddNewContact(driver, extentTest, testName, testDataIteration);
                generic.VerifyContactField(driver, extentTest, testName, testDataIteration, Ro["Lastname1"].ToString(), Ro["Email1"].ToString(), Ro["BussinessPhone"].ToString(), Ro["Language"].ToString(),now);
                generic.Logout(driver, extentTest, testName, testDataIteration);
            }



            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }
        }

    }
}
