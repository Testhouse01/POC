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
using System.Collections.Generic;
using System.Data;

namespace SeleniumCSharpMSTest.TestScripts
{
    [TestClass]

    public class Activities : BaseClass
    {

        LoginFunctions login = new LoginFunctions();
        GenericFunctions generic = new GenericFunctions();

        // <summary>
        /// To verify that the user should be able to fill Call Outcome field as "No services required" and then mark the phone call as complete
        /// </summary>
        [TestCategory("smoke"), TestCategory("IT Sales"), TestCategory("RerunMay120"), TestCategory("Regression1")]
        [TestProperty("TestcaseID", "RTA-5308")]
        public static IEnumerable<object[]> RTA_5848_Data()
        {
            foreach (DataRow row in getTestCaseList("ITSalesManager"))
            {
                yield return new object[] { row };
            }
        }
        [DataTestMethod]
[DynamicData(nameof(RTA_5848_Data), DynamicDataSourceType.Method)]
        public void RTA5308_UsershouldbeabletofilLCallOutcomefieldasNoservicesrequiredandthenmarkthephonecallascomplete(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open IT Sales Opportunities");
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
                string[] s = generic.CreateandVerifyPhoneCall(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), Ro["Button"].ToString(), Ro["Activity"].ToString(), Ro["Call To"].ToString(), Ro["Quick Create Form Button"].ToString(), Ro["Call Outcome1"].ToString(), Ro["Action"].ToString());
                generic.MarkComplete(driver, extentTest, testName, testDataIteration, Ro["Action"].ToString());
                generic.VerifyPhoneCallActivityisReadOnlyOnMarkComplete(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), s[0]);
                login.Logout(driver, extentTest, testName, testDataIteration);
            }

            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }

        }
        public static IEnumerable<object[]> Test12()
        {
            foreach (DataRow row in getTestCaseList("ITSalesManager"))
            {
                yield return new object[] { row };
            }
        }

        // <summary>
        /// To verify that the user should be able to fill Call Outcome field as "Reached IT" and then mark the phone call as complete
        /// </summary>
        [TestCategory("smoke"), TestCategory("IT Sales"), TestCategory("RerunMay120"), TestCategory("Regression1"), TestCategory("RF24-4-20")]
        [TestProperty("TestcaseID", "RTA-5299")]

        [DataTestMethod]
[DynamicData(nameof(Test12), DynamicDataSourceType.Method)]
        public void UsershouldbeabletofilLCallOutcomefieldasReachedITandthenmarkthephonecallascomplete(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open IT Sales Opportunities");
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
                string[] s = generic.CreateandVerifyPhoneCall(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), Ro["Button"].ToString(), Ro["Activity"].ToString(), Ro["Call To"].ToString(), Ro["Quick Create Form Button"].ToString(), Ro["Call Outcome2"].ToString(), Ro["Action"].ToString());
                generic.MarkComplete(driver, extentTest, testName, testDataIteration, Ro["Action"].ToString());
                generic.VerifyPhoneCallActivityisReadOnlyOnMarkComplete(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), s[0]);
                login.Logout(driver, extentTest, testName, testDataIteration);
            }

            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }

        }
        public static IEnumerable<object[]> RTA5311()
        {
            foreach (DataRow row in getTestCaseList("ITSalesManager"))
            {
                yield return new object[] { row };
            }
        }
        // <summary>
        /// To verify that the user should be able to fill Call Outcome field as "Further action scheduled" and then mark the phone call as complete
        /// </summary>
        [TestCategory("smoke"), TestCategory("IT Sales"), TestCategory("RerunMay120"), TestCategory("Regression1")]
        [TestProperty("TestcaseID", "RTA-5311")]

        [DataTestMethod]
[DynamicData(nameof(RTA5311), DynamicDataSourceType.Method)]
        public void RTA5311_UsershouldbeabletofillCallOutcomefieldasCustomerNotReachedandthenmarkthephonecallascomplete(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open IT Sales Opportunities");
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
                string[] s = generic.CreateandVerifyPhoneCall(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), Ro["Button"].ToString(), Ro["Activity"].ToString(), Ro["Call To"].ToString(), Ro["Quick Create Form Button"].ToString(), Ro["Call Outcome3"].ToString(), Ro["Action"].ToString());
                generic.MarkComplete(driver, extentTest, testName, testDataIteration, Ro["Action"].ToString());
                generic.VerifyPhoneCallActivityisReadOnlyOnMarkComplete(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), s[0]);
                login.Logout(driver, extentTest, testName, testDataIteration);
            }

            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }

        }
        public static IEnumerable<object[]> Test111()
        {
            foreach (DataRow row in getTestCaseList("ITSalesManager"))
            {
                yield return new object[] { row };
            }
        }
        // <summary>
        /// To verify that the user should be able to fill Call Outcome field as "Further action scheduled" and then mark the phone call as complete
        /// </summary>
        [TestCategory("smoke"), TestCategory("IT Sales"), TestCategory("RerunMay120"), TestCategory("Regression1")]
        [TestProperty("TestcaseID", "RTA-5300")]

        [DataTestMethod]
[DynamicData(nameof(Test111), DynamicDataSourceType.Method)]
        public void UsershouldbeabletofillCallOutcomefieldasFurtheractionscheduledandthenmarkthephonecallascomplete(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open IT Sales Opportunities");
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
                string[] s = generic.CreateandVerifyPhoneCall(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), Ro["Button"].ToString(), Ro["Activity"].ToString(), Ro["Call To"].ToString(), Ro["Quick Create Form Button"].ToString(), Ro["Call Outcome4"].ToString(), Ro["Action"].ToString());
                generic.MarkComplete(driver, extentTest, testName, testDataIteration, Ro["Action"].ToString());
                generic.VerifyPhoneCallActivityisReadOnlyOnMarkComplete(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), s[0]);
                login.Logout(driver, extentTest, testName, testDataIteration);

            }

            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }

        }
        public static IEnumerable<object[]> Test112()
        {
            foreach (DataRow row in getTestCaseList("ITSalesManager"))
            {
                yield return new object[] { row };
            }
        }
        // <summary>
        /// To verify that the user should be able to fill Call Outcome field as "Wrong Contact Details" and then mark the phone call as complete
        /// </summary>
        [TestCategory("smoke"), TestCategory("IT Sales"), TestCategory("RerunMay120"), TestCategory("Regression1"), TestCategory("RF24-4-20")]
        [TestProperty("TestcaseID", "RTA-5317")]

        [DataTestMethod]
[DynamicData(nameof(Test112), DynamicDataSourceType.Method)]
        public void UsershouldbeabletofillCallOutcomefieldasWrongcontactdetailsandthenmarkthephonecallascomplete(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open IT Sales Opportunities");
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
                string[] s = generic.CreateandVerifyPhoneCall(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), Ro["Button"].ToString(), Ro["Activity"].ToString(), Ro["Call To"].ToString(), Ro["Quick Create Form Button"].ToString(), Ro["Call Outcome5"].ToString(), Ro["Action"].ToString());
                generic.MarkComplete(driver, extentTest, testName, testDataIteration, Ro["Action"].ToString());
                generic.VerifyPhoneCallActivityisReadOnlyOnMarkComplete(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), s[0]);
                login.Logout(driver, extentTest, testName, testDataIteration);

            }

            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }

        }
        public static IEnumerable<object[]> Test113()
        {
            foreach (DataRow row in getTestCaseList("ITSalesManager"))
            {
                yield return new object[] { row };
            }
        }
        // <summary>
        /// To verify that the system will schedule a new phone call with customer based on the newly scheduled date
        /// </summary>
        [TestCategory("smoke"), TestCategory("IT Sales"), TestCategory("RerunMay120"), TestCategory("Regression1")]
        [TestProperty("TestcaseID", "RTA-5306")]

        [DataTestMethod]
[DynamicData(nameof(Test113), DynamicDataSourceType.Method)]
        public void Verifythatthesystemwillscheduleanewphonecallwithcustomerbasedonthenewlyscheduleddate(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open IT Sales Opportunities");
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
                String[] s = generic.CreateandVerifyPhoneCall(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), Ro["Button"].ToString(), Ro["Activity"].ToString(), Ro["Call To"].ToString(), Ro["Quick Create Form Button"].ToString(), Ro["Call Outcome4"].ToString(), Ro["Action"].ToString());
                generic.MarkComplete(driver, extentTest, testName, testDataIteration, Ro["Action"].ToString());
                generic.VerifyPhoneCallActivityisReadOnlyOnMarkComplete(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), s[0]);
                generic.VerifywhethernewPhonecallActivityisCreated(driver, extentTest, testName, testDataIteration, s[0], s[1], Ro["Entity2"].ToString());
                login.Logout(driver, extentTest, testName, testDataIteration);

            }

            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }

        }

        public static IEnumerable<object[]> Test114()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
        // <summary>
        /// To Verify that IT Sales user can Create a phone call from the opportunity and the Sales pipeline in the phone call will be filled with the same value as the opportunity.
        /// </summary>
        [TestCategory("smoke"), TestCategory("IT Sales"), TestCategory("RerunMay120"), TestCategory("Regression1")]
        [TestProperty("TestcaseID", "RTA-5292")]

        [DataTestMethod]
[DynamicData(nameof(Test114), DynamicDataSourceType.Method)]
        public void ITSalesusercanCreateaphonecallfromtheopportunityandtheSalespipelineinthephonecallwillbefilledwiththesamevalue(DataRow Ro)
        {
            try
            {

                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open IT Sales Opportunities");
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                string stage = generic.VerifyBPFinOpportunityScreen(driver, extentTest, testName, testDataIteration);
                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
                generic.CreateandVerifyPhoneCall(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), Ro["Button"].ToString(), Ro["Activity"].ToString(), Ro["Call To"].ToString(), Ro["Quick Create Form Button"].ToString(), Ro["Call Outcome5"].ToString(), Ro["Action"].ToString());
                generic.VerifySalesPipeline(driver, extentTest, testName, testDataIteration, stage);
                login.Logout(driver, extentTest, testName, testDataIteration);
            }

            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }

        }
        public static IEnumerable<object[]> Test115()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
        // <summary>
        /// IT Sales user needs to Verify Newly created Phone call and their Fields
        /// </summary>
        [TestCategory("smoke"), TestCategory("IT Sales"), TestCategory("RerunMay120"), TestCategory("Regression1")]
        [TestProperty("TestcaseID", "RTA-5293")]

        [DataTestMethod]
[DynamicData(nameof(Test115), DynamicDataSourceType.Method)]
        public void VerifyNewlycreatedPhonecallandtheirFields(DataRow Ro)
        {
            try
            {

                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open IT Sales Opportunities");
                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
                generic.QuickCreatePhonecall(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString(), Ro["Activity"].ToString(), Ro["Call To"].ToString(), Ro["Quick Create Form Button"].ToString());
                // generic.VerifyPhoneCallFieldsinTable(driver, extentTest, testName, testDataIteration, Ro["Column1"].ToString(), Ro["Column2"].ToString(), Ro["Column3"].ToString(), Ro["Column4"].ToString(), Ro["Column5"].ToString(), Ro["Column6"].ToString(), Ro["Column7"].ToString(), Ro["Column8"].ToString(), Ro["Column9"].ToString());
                login.Logout(driver, extentTest, testName, testDataIteration);
            }

            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
                Assert.Fail(e.Message);
            }

        }


//        public static IEnumerable<object[]> RTA5319()
//        {
//            foreach (DataRow row in getTestCaseList("ITSalesManager"))
//            {
//                yield return new object[] { row };
//            }
//        }
//        ///To verify that the system will create an open email activity asking to complete customer data if phone call outcome is "Wrong contact details"	
//        /// </summary>
//        [TestCategory("smoke"), TestCategory("IT Sales"), TestCategory("RerunMay120"), TestCategory("Regression1")]
//        [TestProperty("TestcaseID", "RTA-5319")]

//        [DataTestMethod]
//[DynamicData(nameof(RTA5319), DynamicDataSourceType.Method)]
//        public void RTA5319_Systemwillcreateanopenemailactivity(DataRow Ro)
//        {
//            try
//            {
//                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
//                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
//                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
//                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open IT Sales Opportunities");
//                string opp = generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
//                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
//                string[] sub = generic.CreateandVerifyPhoneCall(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString(), Ro["Button"].ToString(), Ro["Activity"].ToString(), Ro["Call To"].ToString(), Ro["Quick Create Form Button"].ToString(), Ro["Call Outcome5"].ToString(), Ro["Action"].ToString());
//                generic.MarkComplete(driver, extentTest, testName, testDataIteration, Ro["Action"].ToString());
//                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
//                generic.SearchforRecord(driver, extentTest, testName, testDataIteration, opp);
//                generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
//                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
//                // generic.VerifywhethernewEmailActivityisCreated(driver, extentTest, testName, testDataIteration, Ro["Email Sub"].ToString(), sub[2]);

//                generic.VerifywhethernewEmailActivityisCreated(driver, extentTest, testName, testDataIteration, Ro["Email Sub"].ToString());

//                login.Logout(driver, extentTest, testName, testDataIteration);

//            }

//            catch (Exception e)
//            {
//                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "The testcase failed due to the following error :  " + e, " ");
//                Assert.Fail(e.Message);
//            }


//        }

    }
}
