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

    public class Contacts:BaseClass
    {
        GenericFunctions generic = new GenericFunctions();
        LoginFunctions login = new LoginFunctions();

        /// <summary>
        /// To verify whetehr IT Sales manager is able to update existing contact record
        /// </summary>
        [TestCategory("smoke"), TestCategory("RerunMay120"),TestCategory("IT Sales"), TestCategory("Regression1")]
        [TestProperty("TestcaseID", "RTA-5853")]
        public static IEnumerable<object[]> RTA_5848_Data()
        {
            foreach (DataRow row in getTestCaseList("ITSalesManager"))
            {
                yield return new object[] { row };
            }
        }
[DataTestMethod]
[DynamicData(nameof(RTA_5848_Data), DynamicDataSourceType.Method)]
        public void RTA5853_ToverifyanITSalesManagerisabletoupdateacontactrecord(DataRow Ro)
        {
            try
            {

                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.UpdateandVerifyContacts(driver, extentTest, testName, testDataIteration, Ro["Field"].ToString(), Ro["Name"].ToString(), Ro["Language Preference"].ToString());
                login.Logout(driver, extentTest, testName, testDataIteration);
            }


            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);

            }

        }

        public static IEnumerable<object[]> Test124()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To verify that account can be linked to contact in IT sales opportunity if they are not linked.
        /// </summary>
        [TestCategory("smoke"), TestCategory("InProgress"), TestCategory("NoRework")]
        [TestProperty("TestcaseID", "RTA-6133")]
        
[DataTestMethod]
[DynamicData(nameof(Test124), DynamicDataSourceType.Method)]
        public void ITSalesNewContactAccountVerification(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectNewTab(driver, extentTest, testName, testDataIteration);
                string TimeInSec = generic.GetSystemTimeInSec(driver, extentTest, testName, testDataIteration);
                generic.AddNewContactAccountinOpportunity(driver, extentTest, testName, testDataIteration, Ro["Topic"].ToString(), TimeInSec, Ro["LName"].ToString(), Ro["Language"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.DiscardChanges(driver, extentTest, testName, testDataIteration);
                ThinkTime(3);
                generic.SelectNewTab(driver, extentTest, testName, testDataIteration);
                generic.AddNewContactITSales(driver, extentTest, testName, testDataIteration, Ro["LName"].ToString(), Ro["AccountName"].ToString(), Ro["Language"].ToString(), TimeInSec);
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectNewTab(driver, extentTest, testName, testDataIteration);
                generic.VerifyContactAndAccount(driver, extentTest, testName, testDataIteration, Ro["Topic"].ToString(), Ro["LName"].ToString(), TimeInSec, Ro["AccountName"].ToString());
                generic.Logout(driver, extentTest, testName, testDataIteration);
            }



            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }
        }

        public static IEnumerable<object[]> Test125()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To verify that account field automatically gets filled in IT sales opportunity when contact is added.
        /// </summary>
        [TestCategory("smoke"), TestCategory("InProgress"), TestCategory("NoRework")]
        [TestProperty("TestcaseID", "RTA-6129")]
       
[DataTestMethod]
[DynamicData(nameof(Test125), DynamicDataSourceType.Method)]
        public void ITSalesExistingContactAccountVerification(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.SelectNewTab(driver, extentTest, testName, testDataIteration);
                string TimeInSec = generic.GetSystemTimeInSec(driver, extentTest, testName, testDataIteration);
                generic.AddNewContactITSales(driver, extentTest, testName, testDataIteration, Ro["LName"].ToString(), Ro["AccountName"].ToString(), Ro["Language"].ToString(), TimeInSec);
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectNewTab(driver, extentTest, testName, testDataIteration);
                generic.VerifyContactAndAccount(driver, extentTest, testName, testDataIteration, Ro["Topic"].ToString(), Ro["LName"].ToString(), TimeInSec, Ro["AccountName"].ToString());
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
