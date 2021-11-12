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

    public class Dashboard :BaseClass
    {
        LoginFunctions login = new LoginFunctions();
        GenericFunctions generic = new GenericFunctions();

        /// <summary>
        /// To check whether Enterprise Sales Manager is able to access Excel templates for opportunities won this month.
        /// </summary>
        [TestCategory("smoke"), TestCategory("RerunMay120"),TestCategory("RerunMay120"), TestCategory("Enterprise Sales"), TestCategory("Regression1")]
        [TestProperty("TestcaseID", "RTA-5321")]
        public static IEnumerable<object[]> RTA_5848_Data()
        {
            foreach (DataRow row in getTestCaseList("EnterpriseSalesManager"))
            {
                yield return new object[] { row };
            }
        }
[DataTestMethod]
[DynamicData(nameof(RTA_5848_Data), DynamicDataSourceType.Method)]

        public void EnterpriseSalesManagerisabletoaccessExceltemplatesforopportunitieswonthismonth(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open Enterprise opportunities");
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Close Button"].ToString());
                generic.CloseasLostWon(driver, extentTest, testName, testDataIteration, Ro["Won"].ToString(), Ro["Status Reason Won"].ToString(), Ro["Closelostwonpopup"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Enterprise Deals Won This Month");
                generic.SelectHeaderButton(driver, extentTest, testName, testDataIteration, Ro["Header Button"].ToString());
                generic.SelectandVerifyExcelTemplate(driver, extentTest, testName, testDataIteration, Ro["Excel Template Option1"].ToString(), Ro["Excel Template Option2"].ToString(), Ro["Close or Return"].ToString());
                login.Logout(driver, extentTest, testName, testDataIteration);
            }

            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }

        }

        public static IEnumerable<object[]> Test126()
        {
            foreach (DataRow row in getTestCaseList("EnterpriseSalesManager"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To check whether Enterprise Sales Manager is able to access Excel templates for opportunities won.
        /// </summary>
        [TestCategory("smoke"), TestCategory("Enterprise Sales"), TestCategory("Regression1")]
        [TestProperty("TestcaseID", "RTA-5323")]
        
[DataTestMethod]
[DynamicData(nameof(Test126), DynamicDataSourceType.Method)]

        public void EnterpriseSalesManagerisabletoaccessExceltemplatesforopportunitieswon(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Enterprise Opportunities Won");
                generic.SelectHeaderButton(driver, extentTest, testName, testDataIteration, Ro["Header Button"].ToString());
                generic.SelectandVerifyExcelTemplate(driver, extentTest, testName, testDataIteration, Ro["Excel Template Option1"].ToString(), Ro["Excel Template Option2"].ToString(), Ro["Close or Return"].ToString());
                login.Logout(driver, extentTest, testName, testDataIteration);
            }

            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }

        }

    }
}
