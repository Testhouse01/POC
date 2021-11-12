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

    public class Opportunity:BaseClass
    {
        LoginFunctions login = new LoginFunctions();
        GenericFunctions generic = new GenericFunctions();



        /// <summary>
        /// To verify whether IT Sales user is able to see the activities
        /// </summary>
        [TestCategory("smoke"), TestCategory("IT Sales"), TestCategory("RerunMay120"), TestCategory("Regression1")]
        [TestProperty("TestcaseID", "RTA-5553")]
        public static IEnumerable<object[]> RTA_5848_Data()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
[DataTestMethod]
[DynamicData(nameof(RTA_5848_Data), DynamicDataSourceType.Method)]
        public void VerifythattheITSalesUserisabletoseetheOpportunityActivitiesTabinOpportunityform(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open IT Sales Opportunities");
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
                generic.VerifyActivities(driver, extentTest, testName, testDataIteration);
                generic.Logout(driver, extentTest, testName, testDataIteration);
            }
            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }
        }

        public static IEnumerable<object[]> Test127()
        {
            foreach (DataRow row in getTestCaseList("ITSalesManager"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// Test Method to "Login to CRM" - Example on how rerun of test method can be Implemented if a element identification fails
        /// </summary>
        [TestCategory("smoke"), TestCategory("InProgress"), TestCategory("NoRework")]
        [TestProperty("TestcaseID", "RTA-9571")]
        
[DataTestMethod]
[DynamicData(nameof(Test127), DynamicDataSourceType.Method)]
        public void ITSalesManagerEditOpportunity(DataRow Ro)
        {
            try
            {

                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, Ro["FilterRecords"].ToString());
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.SelectAndVerifyCustomerFieldITSales(driver, extentTest, testName, testDataIteration, Ro["Topic"].ToString(), Ro["Customer"].ToString(), Ro["Country"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.SelectNewTab(driver, extentTest, testName, testDataIteration);
                generic.SelectAndVerifyCustomerFieldITSales(driver, extentTest, testName, testDataIteration, Ro["Topic"].ToString(), Ro["Customer"].ToString(), Ro["Country"].ToString());
                generic.Logout(driver, extentTest, testName, testDataIteration);
            }



            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }
        }
        public static IEnumerable<object[]> Test128()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        [TestCategory("smoke"), TestCategory("Sales"), TestCategory("RerunMay120"), TestCategory("Regression1"), TestCategory("RF24-4-20")]
        [TestProperty("TestcaseID", "RTA-5028")]
        
[DataTestMethod]
[DynamicData(nameof(Test128), DynamicDataSourceType.Method)]
        public void verifythesystemwilldisplayautopopulatedCentrevaluesameasRecommendedBusinessCentrefieldvalueintheOpportunityrecord(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open Opportunities");
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.HidingPureCloud(driver, extentTest, testName, testDataIteration);
                generic.SelectingandVerifyingRecommendedBusinessCentre(driver, extentTest, testName, testDataIteration, Ro["Centre"].ToString(), Ro["Button"].ToString());
                login.Logout(driver, extentTest, testName, testDataIteration);
            }

            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }
        }

        public static IEnumerable<object[]> RTA5309()
        {
            foreach (DataRow row in getTestCaseList("ITSalesManager"))
            {
                yield return new object[] { row };
            }
        }
        // <summary>
        ///To verify that the user is able to close the opportunity as Close as lost if the phone call outcome is No services required
        /// </summary>
        [TestCategory("smoke"), TestCategory("IT Sales"), TestCategory("RerunMay120"),TestCategory("Regression1"), TestCategory("RF24-4-20")]
        [TestProperty("TestcaseID", "RTA-5309")]
        
[DataTestMethod]
[DynamicData(nameof(RTA5309), DynamicDataSourceType.Method)]
        public void RTA5309_VerifyUserisabletoCloseOpportunityasLost(DataRow Ro)
        {
            try
            {

                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open IT Sales Opportunities");
                string s = generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
                String[] s1 = generic.CreateandVerifyPhoneCall(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), Ro["Button"].ToString(), Ro["Activity"].ToString(), Ro["Call To"].ToString(), Ro["Quick Create Form Button"].ToString(), Ro["Call Outcome1"].ToString(), Ro["Action"].ToString());
                generic.MarkComplete(driver, extentTest, testName, testDataIteration, Ro["Action"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Closed Opportunities");
                generic.SearchforRecord(driver, extentTest, testName, testDataIteration, s);
                generic.ClickParticularRecord(driver, extentTest, testName, testDataIteration, s);
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["LostButton"].ToString());
                generic.CloseasLostWon(driver, extentTest, testName, testDataIteration, Ro["Lost"].ToString(), Ro["Status Reason Lost"].ToString(), Ro["Closelostwonpopup"].ToString());
                login.Logout(driver, extentTest, testName, testDataIteration);

            }

            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }

        }

        public static IEnumerable<object[]> RTA_5852()
        {
            foreach (DataRow row in getTestCaseList("ITSalesManager"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To verify an IT Sales Manager is able to update account record
        /// </summary>
        [TestCategory("smoke"), TestCategory("IT Sales"), TestCategory("RerunMay120"),TestCategory("Regression1")]
        [TestProperty("TestcaseID", "RTA-5852")]
        
[DataTestMethod]
[DynamicData(nameof(RTA_5852), DynamicDataSourceType.Method)]

        public void RTA_5852_ITSalesManagerisabletoupdaterecord(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity3"].ToString());
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.UpdateandVerifyAccounts(driver, extentTest, testName, testDataIteration, Ro["Field2"].ToString(), Ro["Company Name"].ToString());
                login.Logout(driver, extentTest, testName, testDataIteration);

            }

            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }

        }


        
        /// <summary>
        /// To verify that the IT Sales Manager is able to Create, Read and update the Opportunity.


    }
}
