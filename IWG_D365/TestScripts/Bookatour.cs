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
    public class BookaTour : BaseClass
    {

        GenericFunctions generic = new GenericFunctions();
        LoginFunctions login = new LoginFunctions();
        IntegrationFunctions integration = new IntegrationFunctions();
        ExcelReaderFunctions readexcel = new ExcelReaderFunctions();


        public static int maxTestRunsCount = 0;
        public static int maxTestRuns = maxTestRunsCount;
        public static String token;



        /// <summary>///
        /// Method to generate acess token
        /// </summary>
        [ClassInitialize]
        public static void classInitialize(TestContext test)
        {
            IntegrationFunctions integration = new IntegrationFunctions();
            String clientid = "EB31CADD608C4E0CA349DC59C78188C8";
            String password = "825463c42dd50b78b7f283870f3d6649773e1b06d54c9b124b2feb27a4dafe79";
            String tokenjson = integration.accessTokenJSON(clientid, password);
            token = integration.AccessTokenGeneration(tokenjson);

        }


        /// <summary>
        /// Method to start the driver and report
        /// </summary>
        [TestInitialize]
        public void Initialize()
        {
            // var startTime = DateTime.Now.ToString("dd/MM/yyThh:mm:ss");
            TestInitialize();
            maxTestRuns = maxTestRunsCount;
        }

        public void InitializeForRerun()
        {
            TestInitialize();

        }

        public static IEnumerable<object[]> RTA_1()
        {
            foreach (DataRow row in getTestCaseList("DeleteTour"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To verify when an IWG Sales Agent book a tour the slot covers 90 min by default
        /// </summary>

        [TestCategory("smoke"), TestCategory("DeleteallTours")]
        [TestProperty("TestcaseID", "RTA-1")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_1), DynamicDataSourceType.Method)]
        public void Deletebookedtourthisweek(DataRow Ro)
        {

            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());

                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                // login.Logout(driver, extentTest, testName, testDataIteration);

            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }
        }







        public static IEnumerable<object[]> RTA_5005()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To verify when an IWG Sales Agent book a tour the slot covers 90 min by default
        /// </summary>

        [TestCategory("smoke"), TestCategory("RerunMay9"), TestCategory("Validation1"), TestCategory("BAT"), TestCategory("QARefactorfail"), TestCategory("25Aug"), TestCategory("RerunMay120"), TestCategory("30062020_1"), TestCategory("TestLock1"), TestCategory("SprintUN")]
        [TestProperty("TestcaseID", "RTA-5005")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_5005), DynamicDataSourceType.Method)]
        public void RTA_5005_Verifyslotbookedis90minbydefault(DataRow Ro)
        {

            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "My Open Opportunities");


                string parent = driver.CurrentWindowHandle;
                string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["Centre"].ToString());
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
                try
                {
                    string centre = Ro["Centre"].ToString();
                    generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                    generic.CancelTour(driver, extentTest, testName, testDataIteration);



                    driver.SwitchTo().Window(parent);
                    login.Logout(driver, extentTest, testName, testDataIteration);
                }
                catch (Exception e)
                {

                    AddLog(driver, extentTest, testName, testDataIteration, "Info", "Exception handled" + e.Message, "Exception handled");
                }
            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }
        }
        [TestCategory("CompleteTour")]
        [DataTestMethod]
        [DynamicData(nameof(RTA_5005), DynamicDataSourceType.Method)]
        public void CompleteTour(DataRow Ro)
        {

            //try
            //{
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Open Opportunities");
            if (uRL.Contains("uat"))
            {
                generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, "OPP-003321");
            }
            else
            {
                generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, "OPP-020005");
            }
            // generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, "2556");
            // string oppname = generic.GetReferenceIDOpportunity(driver, extentTest, testName, testDataIteration);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            generic.SlotAvailibilityUpdated(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString(), "");


            //login.Logout(driver, extentTest, testName, testDataIteration);

            //}

            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}
        }




        public static IEnumerable<object[]> RTA_5011()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To check IWG Sales Agent is able to see the Centre and Area Sales Manager allocated to the tour booked for the selected centre
        /// </summary>
        [TestCategory("smoke"), TestCategory("RerunMay9"), TestCategory("Validation1"), TestCategory("BAT"), TestCategory("QARefactorfail"), TestCategory("26Aug"), TestCategory("RerunMay120"), TestCategory("Rerun-21062020_1"), TestCategory("Rerun-24062020"), TestCategory("TestLock1"), TestCategory("SprintUN")]
        [TestProperty("TestcaseID", "RTA-5011")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_5011), DynamicDataSourceType.Method)]
        public void RTA_5011_VerifycentreandASMforbookedtour(DataRow Ro)
        {
            try
            {              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                           // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                           // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                           // generic.Logout(driver, extentTest, testName, testDataIteration);
                           // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                //  generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "My Open Opportunities");
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
                string parent = driver.CurrentWindowHandle;

                string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["Centre"].ToString());
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.UpdateStartDate(driver, extentTest, testDataIteration, testName, Ro["SelectMonthNov"].ToString(), Ro["SelectDateNov"].ToString());

                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
                try
                {
                    string centre = Ro["Centre"].ToString();
                    generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                    generic.CancelTour(driver, extentTest, testName, testDataIteration);



                    driver.SwitchTo().Window(parent);
                    login.Logout(driver, extentTest, testName, testDataIteration);
                }
                catch (Exception e)
                {

                    AddLog(driver, extentTest, testName, testDataIteration, "Info", "Exception handled" + e.Message, "Exception handled");
                }
            }

            catch (Exception e)
            {


                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }
        }



        public static IEnumerable<object[]> RTA_5036()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        [Ignore]
        /// <summary>
        /// To check SA is not able to book a tour at blocked time - draganddrop
        /// </summary>

        [TestCategory("RerunMay9"), TestCategory("Sales"), TestCategory("BAT"), TestCategory("Agentissue"), TestCategory("Validation1"), TestCategory("SprintUN")]
        [TestProperty("TestcaseID", "RTA-5036")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_5036), DynamicDataSourceType.Method)]
        public void RTA_5036_Verifytournotbookedforblockeddays(DataRow Ro)
        {
            try
            {              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                           // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                           // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                           // generic.Logout(driver, extentTest, testName, testDataIteration);
                           // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "My Open Opportunities");

                generic.HidingPureCloud(driver, extentTest, testName, testDataIteration);

                string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["BussinessCentre1"].ToString());
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
                string centre = Ro["BussinessCentre1"].ToString();
                generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);

                Assert.Fail("Unable to modify time in the tour confirmation window");

                login.Logout(driver, extentTest, testName, testDataIteration);
            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }
        }



        public static IEnumerable<object[]> RTA_5037()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To verify  that city diary automatically selects the best available resource 
        /// </summary>

        [TestCategory("smoke"), TestCategory("BookaTour"), TestCategory("Validation1"), TestCategory("BAT"), TestCategory("QARefactorfail"), TestCategory("RF24-4-20"), TestCategory("PriorReg3"), TestCategory("TestLock1"), TestCategory("SprintUN"), TestCategory("25aug")]
        [TestProperty("TestcaseID", "RTA-5037")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_5037), DynamicDataSourceType.Method)]
        public void RTA_5037_Verifyslotdisplayasbooked(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "My Open Opportunities");


                string parent = driver.CurrentWindowHandle;
                string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["Centre"].ToString());
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
                try
                {
                    string centre = Ro["Centre"].ToString();
                    generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                    generic.CancelTour(driver, extentTest, testName, testDataIteration);



                    driver.SwitchTo().Window(parent);
                    login.Logout(driver, extentTest, testName, testDataIteration);
                }
                catch (Exception e)
                {

                    AddLog(driver, extentTest, testName, testDataIteration, "Info", "Exception handled" + e.Message, "Exception handled");
                }
            }
            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }
        }



        public static IEnumerable<object[]> RTA_5040()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To verify the SA is not able to book a tour with start date and time in the past
        /// </summary>

        [TestCategory("smoke"), TestCategory("Sales"), TestCategory("Agentissue"), TestCategory("BAT"), TestCategory("25Nov"), TestCategory("26Aug"), TestCategory("Rerun-14-05-2020"), TestCategory("SprintUN"), TestCategory("2910")]
        [TestProperty("TestcaseID", "RTA-5040")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_5040), DynamicDataSourceType.Method)]
        public void RTA_5040_Verifytournotbookedforpastdays(DataRow Ro)
        {
            //try
            //{              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                           // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                           // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                           // generic.Logout(driver, extentTest, testName, testDataIteration);
                           // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "My Open Opportunities");
                string parent = driver.CurrentWindowHandle;
              

                string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["Centre"].ToString());
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
                try
                {
                    string centre = Ro["Centre"].ToString();
                    generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                    generic.CancelTour(driver, extentTest, testName, testDataIteration);



                    driver.SwitchTo().Window(parent);
                    login.Logout(driver, extentTest, testName, testDataIteration);
                //}
                //catch (Exception e)
                //{

                //    AddLog(driver, extentTest, testName, testDataIteration, "Info", "Exception handled" + e.Message, "Exception handled");
                //}
            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }
        }
        public static IEnumerable<object[]> RTA_5060()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }



        public static IEnumerable<object[]> RTA_5044()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To check SA is able to see the same number of concurrent tour for a day as the number of Area Sales manager 
        /// </summary>

        [TestCategory("RerunMay9    "), TestCategory("BookaTour"), TestCategory("Validation1"), TestCategory("BAT"), TestCategory("25Aug"), TestCategory("RerunMay120"), TestCategory("Rerun-25062020"), TestCategory("TestLock1"), TestCategory("SprintUN")]
        [TestProperty("TestcaseID", "RTA-5044")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_5044), DynamicDataSourceType.Method)]
        public void RTA_5044_VerifySAabletoviewtourasASM(DataRow Ro)
        {
            try
            {              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                           // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                           // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                           // generic.Logout(driver, extentTest, testName, testDataIteration);
                           // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());


                string parent = driver.CurrentWindowHandle;
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);

                string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["Centre"].ToString());
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());


                try
                {
                    string centre = Ro["Centre"].ToString();
                    generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                    generic.CancelTour(driver, extentTest, testName, testDataIteration);



                    driver.SwitchTo().Window(parent);
                    login.Logout(driver, extentTest, testName, testDataIteration);
                }
                catch (Exception e)
                {

                    AddLog(driver, extentTest, testName, testDataIteration, "Info", "Exception handled" + e.Message, "Exception handled");
                }
            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }
        }



        public static IEnumerable<object[]> RTA_5058()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To Verify the system should not book the tour if the end time overlaps with another tour booked on the same vertical column
        /// </summary>

        [TestCategory("smoke"), TestCategory("Sales"), TestCategory("Agentissue"), TestCategory("BAT"), TestCategory("Validation1"), TestCategory("25Aug"), TestCategory("Rerun-14-05-2020"), TestCategory("Rerun-15062020"), TestCategory("Rerun-06062020"), TestCategory("Rerun-22062020"), TestCategory("SprintUN"), TestCategory("2910")]
        [TestProperty("TestcaseID", "RTA-5058")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_5058), DynamicDataSourceType.Method)]
        public void RTA_5058_Verifytournotbookedifstarttimeoverlap(DataRow Ro)
        {
            //try
            //{              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                           // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                           // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                           // generic.Logout(driver, extentTest, testName, testDataIteration);
                           // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "My Open Opportunities");
                string parent = driver.CurrentWindowHandle;
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);

                string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["Centre"].ToString());
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
                try
                {
                    string centre = Ro["Centre"].ToString();
                    generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                    generic.CancelTour(driver, extentTest, testName, testDataIteration);



                    driver.SwitchTo().Window(parent);
                    login.Logout(driver, extentTest, testName, testDataIteration);
                }
            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Info", "Exception handled" + e.Message, "Exception handled");
            }
       // }

            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}
        }



        public static IEnumerable<object[]> RTA_5059()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To Verify the system should not book the tour if the end time overlaps with another tour booked on the same vertical column
        /// </summary>

        [TestCategory("smoke"), TestCategory("BookaTour"), TestCategory("Agentissue"), TestCategory("BAT"), TestCategory("25Nov"), TestCategory("25Aug"), TestCategory("Rerun-14-05-2020"), TestCategory("SprintUN")]
        [TestProperty("TestcaseID", "RTA-5059")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_5059), DynamicDataSourceType.Method)]
        public void RTA_5059_Verifytournotbookedifendtimeoverlap(DataRow Ro)
        {
            try
            {              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                           // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                           // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                           // generic.Logout(driver, extentTest, testName, testDataIteration);
                           // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "My Open Opportunities");
                string parent = driver.CurrentWindowHandle;
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);

                string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["Centre"].ToString());
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
                try
                {
                    string centre = Ro["Centre"].ToString();
                    generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                    generic.CancelTour(driver, extentTest, testName, testDataIteration);



                    driver.SwitchTo().Window(parent);
                    login.Logout(driver, extentTest, testName, testDataIteration);
                }
                catch (Exception e)
                {

                    AddLog(driver, extentTest, testName, testDataIteration, "Info", "Exception handled" + e.Message, "Exception handled");
                }
            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }
        }




        public static IEnumerable<object[]> RTA_5063()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To verify each tour vertical slots will display the tour details at the time of mouse hover(click)
        /// </summary>

        [TestCategory("smoke"), TestCategory("BookaTour"), TestCategory("Sales"), TestCategory("BAT"), TestCategory("25Nov"), TestCategory("25Aug"), TestCategory("RerunMay120"), TestCategory("Rerun-06062020"), TestCategory("TestLock1"), TestCategory("SprintUN")]
        [TestProperty("TestcaseID", "RTA-5063")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_5063), DynamicDataSourceType.Method)]
        public void RTA_5063_Verifyverticalslotdisplaydetailsonhover(DataRow Ro)
        {
            try
            {              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                           // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                           // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                           // generic.Logout(driver, extentTest, testName, testDataIteration);
                           // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "My Open Opportunities");
                string parent = driver.CurrentWindowHandle;
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);

                string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["Centre"].ToString());
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());

                try
                {
                    string centre = Ro["Centre"].ToString();
                    generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                    generic.CancelTour(driver, extentTest, testName, testDataIteration);



                    driver.SwitchTo().Window(parent);
                    login.Logout(driver, extentTest, testName, testDataIteration);
                }
                catch (Exception e)
                {

                    AddLog(driver, extentTest, testName, testDataIteration, "Info", "Exception handled" + e.Message, "Exception handled");
                }
            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }
        }



        public static IEnumerable<object[]> RTA_5071()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To check SA is able to book a tour by from calendar view
        /// </summary>


        [TestCategory("RerunMay9"), TestCategory("Sales Agent"), TestCategory("BAT"), TestCategory("BookaTour"), TestCategory("Validation1"), TestCategory("26Aug"), TestCategory("Rerun13-5"), TestCategory("SprintUN")]
        [TestProperty("TestcaseID", "RTA-5071")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_5071), DynamicDataSourceType.Method)]
        public void RTA_5071_SalesAgentBookNewTour(DataRow Ro)
        {
            try
            {              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                           // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                           // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                           // generic.Logout(driver, extentTest, testName, testDataIteration);
                           // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
                string parent = driver.CurrentWindowHandle;
                string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["Centre"].ToString());
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
                try
                {
                    string centre = Ro["Centre"].ToString();
                    generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                    generic.CancelTour(driver, extentTest, testName, testDataIteration);



                    driver.SwitchTo().Window(parent);
                    login.Logout(driver, extentTest, testName, testDataIteration);
                }
                catch (Exception e)
                {

                    AddLog(driver, extentTest, testName, testDataIteration, "Info", "Exception handled" + e.Message, "Exception handled");
                }
            }
            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }
        }





        public static IEnumerable<object[]> RTA_5072()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To verify whether IWG SA is able to change city in the City Diary and select the associated centre
        /// </summary>

        [TestCategory("smoke"), TestCategory("RerunMay9"), TestCategory("2QARefactorfail"), TestCategory("BAT"), TestCategory("Validation1"), TestCategory("28Aug"), TestCategory("RerunMay120"), TestCategory("TestRunJun0806"), TestCategory("21062020_1"), TestCategory("Rerun-24062020"), TestCategory("SprintUN")]
        [TestProperty("TestcaseID", "RTA-5072")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_5072), DynamicDataSourceType.Method)]
        public void RTA_5072_Verifyslotdisplayasbooked(DataRow Ro)
        {

            //try
            //{              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                           // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                           // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                           // generic.Logout(driver, extentTest, testName, testDataIteration);
                           // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "My Open Opportunities");

                // generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
                string parent = driver.CurrentWindowHandle;
                string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["Centre"].ToString());
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.UpdateStartDate(driver, extentTest, testDataIteration, testName, Ro["SelectMonthNov"].ToString(), Ro["SelectDateNov"].ToString());

                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
                try
                {
                    string centre = Ro["Centre"].ToString();
                    generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                    generic.CancelTour(driver, extentTest, testName, testDataIteration);



                    driver.SwitchTo().Window(parent);
                    login.Logout(driver, extentTest, testName, testDataIteration);
                }
                catch (Exception e)
                {

                    AddLog(driver, extentTest, testName, testDataIteration, "Info", "Exception handled" + e.Message, "Exception handled");
                }
            //}
            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}
        }





        public static IEnumerable<object[]> RTA_16949()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        // <summary>
        ///To Verify that when a tour is booked and the current Opp owner is an ASM, the Opp is not reassigned to the new Tour's owner
        //////CRM 5820
        /// </summary>
        [TestCategory("25Nov"), TestCategory("HF"), TestCategory("HF3"), TestCategory("smoke"), TestCategory("BAT"), TestCategory("GlobalP1_S2"), TestCategory("RegressionFail"), TestCategory("11Sep")]
        [TestProperty("TestcaseID", "RTA-16949")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_16949), DynamicDataSourceType.Method)]
        public void RTA_16949_TourBookedOppOwnernotReassignednewtoTourOwner(DataRow Ro)
        {
            try
            {
                string emailid = "crm.test2@regus.com";              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                                                                 // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                                                                 // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                                                                 // generic.Logout(driver, extentTest, testName, testDataIteration);
                                                                 // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            string now = System.DateTime.Now.ToString();
            //string emailid = "crm.test2@regus.com";

            // Create Account
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.SelectNewTab(driver, extentTest, testName, testDataIteration);
            string Brokeraccount = generic.CreateNewAccountWithCustomerTypeNew(driver, extentTest, testDataIteration, testName, Ro["TestAccountLink"].ToString(), now, Ro["CustomerType2"].ToString());
            generic.Movetobroker(driver, extentTest, testName, testDataIteration, Ro["CustomerType2"].ToString());

            //
            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            //  string contactname1 = generic.CreateNewContactEnterpriseSalesForBroker(driver, extentTest, testDataIteration, testName, Ro["Contactname"].ToString(), now, "crm.test4@regus.com", Brokeraccount);
            string contactname = generic.CreateNewContactEnterpriseSalesForBroker(driver, extentTest, testDataIteration, testName, Ro["Contactname"].ToString(), now, emailid, Brokeraccount);

            // Select lead and qualify
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity3"].ToString());
            string now1 = System.DateTime.Now.ToString();

            generic.CreateLeadwithproductITSales(driver, extentTest, testName, testDataIteration, "Lead" + now, Ro["AddEmail"].ToString(), Ro["BussinessPhone"].ToString(), Ro["Source"].ToString(), Ro["MajorSource"].ToString(), Ro["MinorSource2"].ToString(), now1, "ExistingContact", "", contactname, "", "", "Alberta, Calgary - Crowfoot Centre", "", "");
            //generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, "Refresh");
            //generic.QualifyLead1(driver, extentTest, testName, testDataIteration);
            generic.NavigateToRelatedTabEntitiesLead(driver, extentTest, testName, testDataIteration, Ro["RelatedTabEntity"].ToString());
            generic.OppSelect(driver, extentTest, testName, testDataIteration);
            ThinkTime(5);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ThinkTime(6);
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
            // Link Broker Account and contact
            generic.LinkBrokerdetailsNew(driver, extentTest, testName, testDataIteration, "", contactname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            string People = "13";
            generic.AddingNoofPeopleinOpp(driver, extentTest, testName, testDataIteration, People);
            // generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, Ro["salesassistcentre"].ToString(), People);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ThinkTime(10);

            //Assign to Me
            generic.AssignRecordAnotherUser(driver, extentTest, testName, testDataIteration, Ro["AssignTo"].ToString(), Ro["ConnectedTo"].ToString());
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            //Verifying the Opp owner field is now the user
            generic.OwnerOpp(driver, extentTest, testName, testDataIteration, Ro["ConnectedTo"].ToString());
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre"].ToString();

            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ThinkTime(4);
            string TourOwner1 = Element(driver, Control("geTourOwnerID", "CommonObj")).Text;
            //Verifying the Opp owner field is now the tour owner
            // generic.OwnerOpp1(driver, extentTest, testName, testDataIteration, Ro["Owner1"].ToString(),Ro["Owner4"].ToString());
            generic.OwnerOpp(driver, extentTest, testName, testDataIteration, TourOwner1);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());


            //Changing Recommended business centre
            generic.ChangingRecommendedBussinessCentreNew(driver, extentTest, testName, testDataIteration, "London, HQ Moorgate");
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Booking another tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre1 = Ro["Businesscentre1"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ThinkTime(4);
            string TourOwner2 = Element(driver, Control("geTourOwnerID", "CommonObj")).Text;
            //Verifying the Opp owner field is still the first tour owner
            // generic.OwnerOpp1(driver, extentTest, testName, testDataIteration, Ro["Owner1"].ToString(),Ro["Owner4"].ToString());
            generic.OwnerOpp(driver, extentTest, testName, testDataIteration, TourOwner2);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Logout
            login.Logout(driver, extentTest, testName, testDataIteration);

            generic.Loginafterlogout(driver, extentTest, testName, testDataIteration, emailid, Ro["Password1"].ToString());
            generic.NavigateToOutlook(driver, extentTest, testName, testDataIteration);
            generic.SearchAndVerifysystemEmailOutlook(driver, extentTest, testName, testDataIteration, contactname, contactname, "", "Thank you for sending us your referral", "Regus Broker Team");


            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }

        }



        public static IEnumerable<object[]> RTA_16950()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        // <summary>
        ///To Verify that when tour is booked and the current Opp owner is a IWGInternalSales user,the Opp is reassigned to the new Tour's owner
        //////CRM 5820
        /// </summary>
        [TestCategory("HF"), TestCategory("HF3"),TestCategory("smoke"), TestCategory("GlobalP1_S2"), TestCategory("25Nov"), TestCategory("RegressionFail")]
        [TestProperty("TestcaseID", "RTA-16950")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_16950), DynamicDataSourceType.Method)]
        public void RTA_16950_TourBookedOppOwnerReassignedtoTourOwner(DataRow Ro)
        {
            try
            {              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                           // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                           // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                           // generic.Logout(driver, extentTest, testName, testDataIteration);
                           // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            string time = System.DateTime.Now.ToString();

            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            string contactname = generic.CreateNewContactITSalesUser(driver, extentTest, testDataIteration, testName, Ro["contactname"].ToString(), time);

            // Select lead and qualify
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity3"].ToString());
            string now = System.DateTime.Now.ToString();
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.CreateLeadwithproductITSales(driver, extentTest, testName, testDataIteration, "Lead" + now, Ro["AddEmail"].ToString(), Ro["BussinessPhone"].ToString(), Ro["Source"].ToString(), Ro["MajorSource"].ToString(), Ro["MinorSource2"].ToString(), now, "ExistingContact", "", contactname, "", "", "Alberta, Calgary - Crowfoot Centre", "", "");
            //generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, "Refresh");
            //generic.QualifyLead1(driver, extentTest, testName, testDataIteration);
            generic.NavigateToRelatedTabEntitiesLead(driver, extentTest, testName, testDataIteration, Ro["RelatedTabEntity"].ToString());
            generic.OppSelect(driver, extentTest, testName, testDataIteration);
            ThinkTime(5);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ThinkTime(5);
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");


            string People = "13";
            generic.AddingNoofPeopleinOpp(driver, extentTest, testName, testDataIteration, People);
            // generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, Ro["salesassistcentre"].ToString(), People);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ThinkTime(5);

            //Assign to Me
            generic.AssignRecordAnotherUser(driver, extentTest, testName, testDataIteration, Ro["AssignTo"].ToString(), Ro["ConnectedTo"].ToString());
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            //Verifying the Opp owner field is now the user
            generic.OwnerOpp(driver, extentTest, testName, testDataIteration, Ro["ConnectedTo"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            string TourOwner1 = Element(driver, Control("geTourOwnerID", "CommonObj")).Text;
            //Verifying the Opp owner field is now the tour owner
            generic.OwnerOpp(driver, extentTest, testName, testDataIteration, TourOwner1);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            ////Reassign back to loggedin user
            //generic.AssignRecordAnotherUser(driver, extentTest, testName, testDataIteration, Ro["AssignTo"].ToString(), Ro["ConnectedTo"].ToString());
            //generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ////Verifying the Opp owner field is now the user
            //generic.OwnerOpp(driver, extentTest, testName, testDataIteration, Ro["ConnectedTo"].ToString());

           ///// Booking another tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre1 = Ro["Businesscentre1"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            string TourOwner2 = Element(driver, Control("geTourOwnerID", "CommonObj")).Text;
            //Verifying the Opp owner field is now the tour owner
            generic.OwnerOpp(driver, extentTest, testName, testDataIteration, TourOwner2);

            //Logout
            login.Logout(driver, extentTest, testName, testDataIteration);

            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }

        }





        public static IEnumerable<object[]> RTA_17767()
        {
            foreach (DataRow row in getTestCaseList("DirectSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        // <summary>
        ///To Verify whether the user cant close the opportunity as Won,when payment received is No.
        //////CRM 4397
        /// </summary>
        [TestCategory("25Nov1"), TestCategory("CRM-4397"), TestCategory("BAT"), TestCategory("Regression"), TestCategory("GlobalP1_S2"), TestCategory("GR")]
        [TestProperty("TestcaseID", "RTA-17767")]
        [DataTestMethod]
        [DynamicData(nameof(RTA_17767), DynamicDataSourceType.Method)]
        public void RTA_17767_Verify_Dailer_Opp(DataRow Ro)
        {
            //try
            //{              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
            // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
            // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
            // generic.Logout(driver, extentTest, testName, testDataIteration);
            // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            string Time = System.DateTime.Now.ToString();
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            // generic.HeaderButtonwithoutConfirmation(driver, extentTest, testName, testDataIteration, "New");

            generic.CreateOpportunityFormCondition(driver, extentTest, testName, testDataIteration, Time, "true", "true", "true", "true", "true", "true", "true", "true", "true", "true");
            string oppname = generic.GetOppID(driver, extentTest, testName, testDataIteration);

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, "", oppname);

            ThinkTime(5);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, "Dialer - No answer to Post Tour Call");
            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);

            // generic.VerifyDailerPosttour(driver, extentTest, testName, testDataIteration, "Alberta, Calgary - Crowfoot Centre", "Open");
            generic.VerifyDailerPosttourStatus(driver, extentTest, testName, testDataIteration, "Open");

            //Re-schedule tour

            generic.NavigateTourfrompayload(driver, extentTest, testName, testDataIteration);

            ThinkTime(3);
            generic.ScriptErrorExcep(driver, extentTest, testName, testDataIteration);

            generic.RescheduleTour(driver, extentTest, testName, testDataIteration, "");
            generic.ResheduleBookaslot(driver, extentTest, testName, testDataIteration, "", "");

            //Navigating to activities and verifying language in payload
            generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, "Dialer - No answer to Post Tour Call");

            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);

            //  generic.VerifyDailerPosttour(driver, extentTest, testName, testDataIteration, "Alberta, Calgary - Crowfoot Centre", "Open");
            generic.VerifyDailerPosttourStatus(driver, extentTest, testName, testDataIteration, "Open");
            login.Logout(driver, extentTest, testName, testDataIteration);
            //}

            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");
            //}

        }




        public static IEnumerable<object[]> RTA_13922()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }


        /// <summary>
        /// TEST [CRM-5242] Verify that Phone number contained in No Contact Email journey with Brand Regus and languages English or Norwegian delivered to customers must be updated
        /// </summary>
        [Ignore]
        //Marked as ignore as  this fuctionality is been decomisioned as part of global p1 release.
        [TestCategory("Regression"), TestCategory("Priority1"), TestCategory("BAT"), TestCategory("Sprint 39.1"), TestCategory("Sprint3640")]
        [TestProperty("TestcaseID", "RTA-13922")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_13922), DynamicDataSourceType.Method)]
        public void RTA_13922_VerifyPhoneContactonPostTourEmail(DataRow Ro)
        {
            try
            {              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                           // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                           // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                           // generic.Logout(driver, extentTest, testName, testDataIteration);
                           // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "My Open Opportunities");
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);

                string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["BussinessCentre2"].ToString());
                string oppname = generic.GetOppID(driver, extentTest, testDataIteration, testName);
                generic.UpdateStartDate(driver, extentTest, testDataIteration, testName, Ro["SelectMonthNov"].ToString(), Ro["SelectDateNov"].ToString());

                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());

                string centre = Ro["BussinessCentre2"].ToString();
                generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);


                generic.VerifyPhoneisUpdatedonEmails(driver, extentTest, testName, testDataIteration, "(+47) 21 98 48 29");

                login.Logout(driver, extentTest, testName, testDataIteration);
            }


            catch (Exception e)
            {


                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }

        }


        public static IEnumerable<object[]> RTA_17512()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        //TEST[CRM-2799]-Verify whether the tour remainder is set to "yes",24 hours before tour start date
        /// CRM-2799

        /// </summary>

        [TestCategory("Regression"), TestCategory("GlobalP1_S2"), TestCategory("BAT"), TestCategory("vg"), TestCategory("25Nov")]

        [TestProperty("TestcaseID", "RTA-17512")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_17512), DynamicDataSourceType.Method)]
        public void RTA_17512_VerifycommRqstTourRshdld(DataRow Ro)
        {
            //try

            //{              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
            // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
            // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
            // generic.Logout(driver, extentTest, testName, testDataIteration);
            // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);

            string emailid = "AutomatCompany" + Time + "@mailinator.com";
            string centre = "Stuttgart, Konigstrasse";
            string now = System.DateTime.Now.ToString();

            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");



            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            //Creating new Opp
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "6";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, centre, People);

            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

            // Refresh the page
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Verifying Opportunity status and Language

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);

            ThinkTime(5);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["emailconnRqst"].ToString());
            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverification(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), "Regus", emailid, "", "", "", "", "");

            //Re-schedule tour

            generic.NavigateTourfrompayload(driver, extentTest, testName, testDataIteration);

            ThinkTime(3);
            generic.ScriptErrorExcep(driver, extentTest, testName, testDataIteration);


            generic.RescheduleTour(driver, extentTest, testName, testDataIteration, "");
            generic.ResheduleBookaslot(driver, extentTest, testName, testDataIteration, contactname, centre);

            //Navigating to activities and verifying language in payload
            generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["emailconnRqst"].ToString());

            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverification(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), "Regus", emailid, "");
            //CLOSE TOUR
            generic.NavigateTourfrompayload(driver, extentTest, testName, testDataIteration);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Action1"].ToString());

            login.Logout(driver, extentTest, testName, testDataIteration);

            //}
            //catch (Exception e)
            //{
            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");

            //    Type thisType = this.GetType();

            //    object testCall = this;

            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            //}
        }


        public static IEnumerable<object[]> RTA_17661()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        //TEST[CRM-2801]-Verify whether the comm request is created , when the tour is scheduled / rescheduled.
        /// CRM-2801

        /// </summary>

        [TestCategory("25Nov"), TestCategory("Regression"), TestCategory("GlobalP1_S2"), TestCategory("28Sep2020"), TestCategory("1410"), TestCategory("27Aug"), TestCategory("vg"), TestCategory("2810"), TestCategory("2910")]


        [TestProperty("TestcaseID", "RTA-17661")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_17661), DynamicDataSourceType.Method)]
        public void RTA_17661_VerifycommRqstTourRshdld(DataRow Ro)
        {
            //try

            //{              
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);

            string emailid = "AutomatCompany" + Time + "@mailinator.com";
            string centre = "London, HQ Moorgate ";
            string now = System.DateTime.Now.ToString();

            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");


            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            //Creating new Opp
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "6";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, centre, People);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ThinkTime(2);

            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");



            //// Refresh the page
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);

            //ThinkTime(5);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            string parent = driver.CurrentWindowHandle;
            generic.AdvancedFind(driver, extentTest, testName, testDataIteration);

            // Enter filter criteria
            generic.advancefindfiltertourinopp(driver, extentTest, testName, testDataIteration, "Tours", "Opportunity", "Equals", oppname);
            generic.VerifyingTourRemainder(driver, extentTest, testName, testDataIteration, "Yes");
            driver.Close();
            driver.SwitchTo().Window(parent);

            //CLOSE TOUR
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
            generic.OpenExistingTour(driver, extentTest, testName, testDataIteration);
            ThinkTime(3);


            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Action1"].ToString());



            login.Logout(driver, extentTest, testName, testDataIteration);

            //}
            //catch (Exception e)
            //{
            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");

            //    Type thisType = this.GetType();

            //    object testCall = this;

            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            //}
        }


        public static IEnumerable<object[]> RTA_17687()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        ///TEST[CRM-3026] Verify that No answer to Post Tour Call - No Show - Reschedule your visit" Comm Request is not sent when Opportunity contact has 'Do Not Contact' field = 'Yes'.
        /// CRM-3026, 7559, 7859
        /// </summary>
        [TestCategory("0409"), TestCategory("27Aug"), TestCategory("HF"), TestCategory("Regression"), TestCategory("GlobalP1_S2"), TestCategory("HF3"), TestCategory("2810")]
        [TestProperty("TestcaseID", "RTA-17687")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_17687), DynamicDataSourceType.Method)]
        public void RTA_17687_CommReqstposttourattendopp(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@gmail.com";

            //Navigate to Opportunity to verify closed and lost Opp count is same as adv find result
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "IWG Direct Agent Performance dashboard");

            //Closed Opportunities                
            generic.VerifyOppcount(driver, extentTest, testName, testDataIteration, Ro["ResultsButton"].ToString());

             //Create new Opportunity for Book a tour
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            // generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            string contactname = generic.CreateNewcontactOpportunity(driver, extentTest, testName, testDataIteration, Ro["BussinessCentre2"].ToString(), Ro["Language1"].ToString(), emailid);
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");


            //Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            /// changing contact preference
            generic.contactpreferencesettoYes(driver, extentTest, testName, testDataIteration, "Details");

            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre1"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //// //Verifying two timezone fields in tour record

            //Selecting Activities and filling values for tour outcome
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
            generic.TourTimeZone(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle1"].ToString(), Ro["YourZone"].ToString(), Ro["CentreZone"].ToString());

            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);


            //Selecting Activities and filling values for tour outcome
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());

            generic.TourOutcomeArrive(driver, extentTest, testName, testDataIteration, Ro["CreatedTour"].ToString(), Ro["P1noshow"].ToString());
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Navigating to activities not present
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.selectNoAnswerTourRqst(driver, extentTest, testName, testDataIteration, Ro["Psttournocontact"].ToString());


            //Verifying Payload response Request not processed
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverifyrqstnotprocessed(driver, extentTest, testName, testDataIteration, Ro["ContactRqstnotProcessed"].ToString(), Ro["Psttournocontact"].ToString());

            login.Logout(driver, extentTest, testName, testDataIteration);
            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }

        }



        public static IEnumerable<object[]> RTA_16472()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        /// <summary>
        //TEST [CRM-3218] Verify that correct signature phone number and email address are sent in the "No answer to Post Tour Call - No Show" Comm Request

        /// CRM-3218
        /// </summary>
        [TestCategory("25Nov"), TestCategory("GlobalP1_S2"), TestCategory("BAT"), TestCategory("vg")]
        [TestProperty("TestcaseID", "RTA-16472")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_16472), DynamicDataSourceType.Method)]
        public void RTA_16472_CommReqstposttourattendopp(DataRow Ro)
        {

            //try
            //{             
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);

            string centre = "Oslo, Spaces Nydalen";
            string emailid = "AutomatCompany" + Time + "@gmail.com";

            string now = System.DateTime.Now.ToString();


            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");
            generic.EnterContactNumber(driver, extentTest, testName, testDataIteration, Ro["BussinessPhone"].ToString());


            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            //Creating new Opp
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "6";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, centre, People);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ThinkTime(2);
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
            //Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre1 = Ro["BussinessCentre1"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());



            //Selecting Activities and filling values for tour outcome
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());

            //generic.SelectActiveCellContacts(driver, extentTest, testName, testDataIteration);

            generic.TourOutcomeArrive(driver, extentTest, testName, testDataIteration, Ro["CreatedTour"].ToString(), Ro["P1noshow"].ToString());
            //generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);


            //Navigating to activities not present
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["Psttournocontact"].ToString());

            //Verifying Payload response Request not processed
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverificationwithphone(driver, extentTest, testName, testDataIteration, Ro["Languagecode1"].ToString(), "Spaces", emailid, Ro["payloadphone1"].ToString(), Ro["TeamEmailqa"].ToString());




            login.Logout(driver, extentTest, testName, testDataIteration);
            //    }

            //        catch (Exception e)
            //        {

            //            AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}

        }


        public static IEnumerable<object[]> RTA_16474()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        ///TEST [CRM-3218] Verify that correct signature phone number and email address are sent in the "ASM, AD Internal Handover Email - Large Deal Tour" Comm Request
        /// </summary>
        /// CRM-3218
        [TestCategory("HF"), TestCategory("GlobalP1_S2"), TestCategory("HF3"), TestCategory("BAT"), TestCategory("25Nov")]

        [TestProperty("TestcaseID", "RTA-16474")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_16474), DynamicDataSourceType.Method)]

        public void RTA_16474_Verifycommuequestgeneratedtourscheduled(DataRow Ro)
        {

            try

            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);

            string centre = "London, Goodge Street";
            string emailid = "AutomatCompany" + Time + "@gmail.com";

            string now = System.DateTime.Now.ToString();


            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");



            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            //Creating new Opp
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "13";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, Ro["BussinessCentre2"].ToString(), People);

            string oppname = generic.GetOppID(driver, extentTest, testName, testDataIteration);




            //Verifying Opportunity status and Language

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());

            string centre1 = Ro["BussinessCentre2"].ToString();

            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());



            //Navigating to activities and verifying language in payload

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["AsmLargeTourReques"].ToString());

            // Verify the owner            
                generic.CommrequestHeaderverification(driver, extentTest, testName, testDataIteration, Ro["Commreqowner"].ToString());
            
            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            // generic.payloadverificationwithphone(driver, extentTest, testName, testDataIteration, Ro["Languagecode"].ToString(), Ro["regusbrand"].ToString(), emailid, Ro["payloadphone"].ToString(), Ro["TeamEmailqa"].ToString());
            generic.payloadverificationwithphone(driver, extentTest, testName, testDataIteration, Ro["Languagecode"].ToString(), Ro["regusbrand"].ToString(), emailid, Ro["phonepayload"].ToString(), Ro["PayloadEmailqa"].ToString());


            login.Logout(driver, extentTest, testName, testDataIteration);

            }

            catch (Exception e)

            {



                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");

                Type thisType = this.GetType();

                object testCall = this;

                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            }

        }




        public static IEnumerable<object[]> RTA_17686()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }


        /// <summary>
        ////Verify whether the comm request"Post tour-no show" is not generated for the Customers by selecting 'Allow Emails' as 'No.
        /// CRM-3026
        /// </summary>
        [TestCategory("0409"), TestCategory("28Aug"), TestCategory("Regression"), TestCategory("BAT"), TestCategory("GlobalP1_S2"), TestCategory("vg")]
        [TestProperty("TestcaseID", "RTA-17686")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_17686), DynamicDataSourceType.Method)]
        public void RTA_17686_CommReqstposttourattendopp(DataRow Ro)
        {
            //try
            //{              
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@gmail.com";

            //Create new Opportunity for Book a tour
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            string contactname = generic.CreateNewcontactOpportunity(driver, extentTest, testName, testDataIteration, Ro["BussinessCentre2"].ToString(), Ro["Language1"].ToString(), emailid);
            string oppname = generic.GetOppID(driver, extentTest, testName, testDataIteration);


            //Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            //generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            /// changing contact preference
            generic.contactpreferenceemailfield(driver, extentTest, testName, testDataIteration, "Details");

            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre1"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());


            //Selecting Activities and filling values for tour outcome
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());

            //generic.SelectActiveCellContacts(driver, extentTest, testName, testDataIteration);

            generic.TourOutcomeArrive(driver, extentTest, testName, testDataIteration, Ro["CreatedTour"].ToString(), Ro["P1noshow"].ToString());
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Navigating to activities not present
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            //generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["Psttournocontact"].ToString());
            generic.selectNoAnswerTourRqst(driver, extentTest, testName, testDataIteration, Ro["Psttournocontact"].ToString());

            //Verifying Payload response Request not processed
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverifyrqstnotprocessed(driver, extentTest, testName, testDataIteration, Ro["Attendednotstatus"].ToString(), Ro["Psttournocontact"].ToString());

            login.Logout(driver, extentTest, testName, testDataIteration);

            //}

            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}

        }





        public static IEnumerable<object[]> RTA_17688()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        ////TEST[CRM - 3026]-Verify that No answer to Post Tour Call - No Show - Reschedule your visit" Comm Request is not sent when Opportunity contact has 'Do Not Contact' field = 'Yes'.
        /// CRM-3026
        /// </summary>
        [TestCategory("0309"), TestCategory("27Aug"), TestCategory("Regression"), TestCategory("BAT"), TestCategory("GlobalP1_S2"), TestCategory("vg")]
        [TestProperty("TestcaseID", "RTA-17688")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_17688), DynamicDataSourceType.Method)]
        public void RTA_17688_CommReqstrposttourattendopp(DataRow Ro)
        {
            //try
            //{             
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@gmail.com";

            //Create new Opportunity for Book a tour//
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            string contactname = generic.CreateNewcontactOpportunity(driver, extentTest, testName, testDataIteration, Ro["BussinessCentre2"].ToString(), Ro["Language1"].ToString(), emailid);
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

            //Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre1"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ///// changing contact preference
            generic.OPPdonotsend(driver, extentTest, testName, testDataIteration, "Yes");


            //Selecting Activities and filling values for tour outcome
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());

            //generic.SelectActiveCellContacts(driver, extentTest, testName, testDataIteration);

            generic.TourOutcomeArrive(driver, extentTest, testName, testDataIteration, Ro["CreatedTour"].ToString(), Ro["P1noshow"].ToString());
            //generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Navigating to activities not present
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            //generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());
            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["Psttournocontact"].ToString());

            //Verifying Payload response Request not processed
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverifyrqstnotprocessed(driver, extentTest, testName, testDataIteration, Ro["Resquestnotprocessed"].ToString(), Ro["Psttournocontact"].ToString());



            login.Logout(driver, extentTest, testName, testDataIteration);
            //}

            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}

        }



        public static IEnumerable<object[]> RTA_17759()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        ////TEST[CRM - 3027]-Verify whether the comm request"No answer to Post Tour Call - Attended" is not generated for the Customers by selecting 'Allow Emails' as 'No.
        /// CRM-3027
        /// </summary>
        [TestCategory("0409"), TestCategory("25Nov"), TestCategory("BAT"), TestCategory("Regression"), TestCategory("GlobalP1_S2"), TestCategory("vg")]
        [TestProperty("TestcaseID", "RTA-17759")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_17759), DynamicDataSourceType.Method)]
        public void RTA_17759_CommReqstposttourattendopp(DataRow Ro)
        {
            //try
            //{              
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@gmail.com";
            String centre = "London, HQ Moorgate";

            //Create new Opportunity for Book a tour
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            string contactname = generic.CreateNewcontactOpportunity(driver, extentTest, testName, testDataIteration, centre, Ro["Language1"].ToString(), emailid);
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

            //Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            /// changing contact preference
            generic.contactpreferenceemailfield(driver, extentTest, testName, testDataIteration, "Details");

            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            //string centre = Ro["BussinessCentre1"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Selecting Activities and filling values for tour outcome
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());


            generic.TourOutcomeArrive(driver, extentTest, testName, testDataIteration, Ro["CreatedTour"].ToString(), Ro["P1show"].ToString());
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Navigating to activities not present
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            //generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());
            generic.selectNoAnswerTourRqst(driver, extentTest, testName, testDataIteration, Ro["posttourDNCrqst"].ToString());

            //Verifying Payload response Request not processed
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverifyrqstnotprocessed(driver, extentTest, testName, testDataIteration, Ro["Attendednotstatus"].ToString(), Ro["posttourDNCrqst"].ToString());



            login.Logout(driver, extentTest, testName, testDataIteration);

            //}
            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}

        }



        public static IEnumerable<object[]> RTA_17765()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        /// <summary>
        ///TEST[CRM-3027]-Verify that No answer to "No answer to Post Tour Call - Attended" Comm Request is not sent when Opportunity has 'Do Not Contact' field = 'Yes'.
        /// CRM-3027
        /// </summary>
        [TestCategory("Regression"), TestCategory("GlobalP1_S2"), TestCategory("BAT"), TestCategory("BAT"), TestCategory("0309"), TestCategory("27Aug"), TestCategory("vg")]
        [TestProperty("TestcaseID", "RTA-17765")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_17765), DynamicDataSourceType.Method)]
        public void RTA_17765_CommReqstposttourattendopp(DataRow Ro)
        {
            //try
            //{              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
            // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
            // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
            // generic.Logout(driver, extentTest, testName, testDataIteration);
            // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@gmail.com";

            //Create new Opportunity for Book a tour//
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            string contactname = generic.CreateNewcontactOpportunity(driver, extentTest, testName, testDataIteration, Ro["BussinessCentre2"].ToString(), Ro["Language1"].ToString(), emailid);
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

            //Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre1"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ///// changing contact preference
            generic.OPPdonotsend(driver, extentTest, testName, testDataIteration, "Yes");


            //Selecting Activities and filling values for tour outcome
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());

            //generic.SelectActiveCellContacts(driver, extentTest, testName, testDataIteration);

            generic.TourOutcomeArrive(driver, extentTest, testName, testDataIteration, Ro["CreatedTour"].ToString(), Ro["P1show"].ToString());
            //generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);


            //Navigating to activities not present
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["posttourDNCrqst"].ToString());

            //Verifying Payload response Request not processed
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverifyrqstnotprocessed(driver, extentTest, testName, testDataIteration, Ro["Resquestnotprocessed"].ToString(), Ro["posttourDNCrqst"].ToString());



            login.Logout(driver, extentTest, testName, testDataIteration);
            //}

            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}

        }




        public static IEnumerable<object[]> RTA_17766()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        /// <summary>
        ///TEST[CRM-2797]-Verify whether the communication request is created for the post tour No show opportunity with brand and language
        /// CRM-2797
        /// </summary>
        [TestCategory("28Sep2020"), TestCategory("Regression"), TestCategory("GlobalP1_S2"), TestCategory("BAT"), TestCategory("vg"), TestCategory("0309"), TestCategory("2810")]
        [TestProperty("TestcaseID", "RTA-17766")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_17766), DynamicDataSourceType.Method)]
        public void RTA_17766_CommReqstforposttourattendedopp(DataRow Ro)
        {

            //try
            //{             
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@gmail.com";

            //Create new Opportunity for Book a tour//
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());

            string contactname = generic.CreateNewcontactOpportunity(driver, extentTest, testName, testDataIteration, Ro["BussinessCentre2"].ToString(), Ro["Language1"].ToString(), emailid);
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            //////Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());


            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre2"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ///////// changing contact preference
            generic.contactpreferencedonotsend(driver, extentTest, testName, testDataIteration, "Details");

            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);


            //Selecting Activities and filling values for tour outcome
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());

            //generic.SelectActiveCellContacts(driver, extentTest, testName, testDataIteration);

            generic.TourOutcomeArrive(driver, extentTest, testName, testDataIteration, Ro["CreatedTour"].ToString(), Ro["P1show"].ToString());
            ////generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);

            //generic.Selectdropdownvalue(driver, extentTest, testName, testDataIteration, "Closed Tours");
            //generic.SelectingActiveAccountEnterpriseSales(driver, extentTest, testName, testDataIteration);

            //generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["Realtedcommrqst"].ToString());
            //generic.Selectdropdownvalue(driver, extentTest, testName, testDataIteration, "Closed Comm Requests");


            ////Navigating to activities and verifying language in payload
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());

            generic.searchrequest(driver, extentTest, testName, testDataIteration, Ro["P1posttournattended"].ToString());

            generic.posttourcommrequestverify(driver, extentTest, testName, testDataIteration, Ro["P1posttournattended"].ToString());


            generic.Selectconnection(driver, extentTest, testName, testDataIteration, Ro["P1posttournattended"].ToString());

            ////Verifying Payload response Request not processed
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverifyrqstnotprocessed(driver, extentTest, testName, testDataIteration, Ro["ContactRqstnotProcessed"].ToString(), Ro["P1posttournattended"].ToString());
            login.Logout(driver, extentTest, testName, testDataIteration);

            //}



            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}

        }



        public static IEnumerable<object[]> RTA_16284()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        /// <summary>
        ///TEST[CRM-2797]-Verify whether the communication request is created for the post tour No show opportunity with brand and language
        /// CRM-2797
        /// </summary>
        [TestCategory("25AugFixed"), TestCategory("Regression"), TestCategory("24Nov"), TestCategory("GlobalP11"), TestCategory("GlobalP11A"), TestCategory("11Sep"), TestCategory("25Sep2020"), TestCategory("2211")]
        [TestProperty("TestcaseID", "RTA-16284")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_16284), DynamicDataSourceType.Method)]
        public void RTA_16284_CommReqstforposttourattendedopp(DataRow Ro)
        {

            //try
            //{              
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@mailinator.com";

            //Create new Opportunity for Book a tour
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            string contactname = generic.CreateNewcontactOpportunityWithBusPh(driver, extentTest, testName, testDataIteration, Ro["BussinessCentre2"].ToString(), Ro["Language1"].ToString(), emailid, Ro["BussinessPhone"].ToString());
            //generic.EnterContactNumber(driver, extentTest, testName, testDataIteration, Ro["BussinessPhone"].ToString());
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");


            // Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre1"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Selecting Activities and filling values for tour outcome
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
            generic.TourOutcome(driver, extentTest, testName, testDataIteration, Ro["CreatedTour"].ToString(), Ro["P1noshow"].ToString(), Ro["P1noshowreason"].ToString());
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());

            //Navigating to activities and verifying language in payload
            generic.searchrequest(driver, extentTest, testName, testDataIteration, Ro["P1posttournoshow"].ToString());
            generic.posttourcommrequestverify(driver, extentTest, testName, testDataIteration, Ro["P1posttournoshow"].ToString());
            generic.SelectingActiverequest(driver, extentTest, testName, testDataIteration);

            //Verifying Payload fields
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverification(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), Ro["regusbrand"].ToString(), emailid,"");

            login.Logout(driver, extentTest, testName, testDataIteration);

            //}

            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            //}

        }




        public static IEnumerable<object[]> RTA_16288()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        /// <summary>

        ///TEST[CRM-2798]Verify whether the communication request is created for the post tour attended  opportunity with brand and language.

        /// CRM-2798

        /// </summary>

        [TestCategory("HF"), TestCategory("HF3") ,TestCategory("Regression"), TestCategory("0911"), TestCategory("GlobalP11"), TestCategory("GlobalP11A"), TestCategory("11Sep"), TestCategory("25Sep2020")]

        [TestProperty("TestcaseID", "RTA-16288")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_16288), DynamicDataSourceType.Method)]

        public void RTA_16288_CommReqstforposttourattendopp(DataRow Ro)
        {

            try
            {
                string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@mailinator.com";
            //string emailid = "test11sep@mailinator.com";

            // Login to the application
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            
            //Create new Opportunity for Book a tour
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            string contactname = generic.CreateNewcontactOpportunity(driver, extentTest, testName, testDataIteration, Ro["BussinessCentre2"].ToString(), Ro["Language1"].ToString(), emailid);
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

            //Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre2"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Selecting Activities and filling values for tour outcome
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
            generic.TourOutcome(driver, extentTest, testName, testDataIteration, Ro["CreatedTour"].ToString(), Ro["P1show"].ToString(), Ro["P1showreason"].ToString());
         
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());

            //Navigating to activities and verifying language in payload
            generic.searchrequest(driver, extentTest, testName, testDataIteration, Ro["P1posttournattended"].ToString());
            generic.posttourcommrequestverify(driver, extentTest, testName, testDataIteration, Ro["P1posttournattended"].ToString());
            generic.SelectingActiverequest(driver, extentTest, testName, testDataIteration);

            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverification(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), Ro["regusbrand"].ToString(), emailid, "");

            // Logout from the application
            login.Logout(driver, extentTest, testName, testDataIteration);


            }
            catch (Exception e)
            {
                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }

        }




        public static IEnumerable<object[]> RTA_16538()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        //Verify whether the communication request is generated for assigned user when the tour is rescheduled.
        /// CRM-2810

        /// </summary>

        [TestCategory("Regression"), TestCategory("GlobalP11_2"), TestCategory("abhi"), TestCategory("BAT"), TestCategory("HF"), TestCategory("GlobalP11A"), TestCategory("2810")]

        [TestProperty("TestcaseID", "RTA-16538")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_16538), DynamicDataSourceType.Method)]
        public void RTA_16538_VerifycommRqstTourRshdld(DataRow Ro)
        {
            //try

            //{             
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);

            string emailid = "AutomatCompany" + Time + "@gmail.com";

            string now = System.DateTime.Now.ToString();

            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");

            //Creating new Opp          

            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "13";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, Ro["BussinessCentre"].ToString(), People);

            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

            ThinkTime(20);
            // Refresh the page
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Navigating to activities and verifying language in payload
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["AsmLargeTourReques"].ToString());

            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            //generic.payloadverificationwithASM(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), "Regus", emailid, Ro["TourASMtest3"].ToString());
            generic.payloadverificationASMBusinessCentre(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), Ro["BussinessCentre"].ToString(), "Regus", emailid, Ro["TourASMtest3"].ToString());

            // Go back to Opportunity
            generic.Backtooppotunity(driver, extentTest, testName, testDataIteration);
            //Re-schedule tour

            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
            generic.OpenExistingTour(driver, extentTest, testName, testDataIteration);
            ThinkTime(3);

            generic.RescheduleTour(driver, extentTest, testName, testDataIteration, "");
            generic.ResheduleBookaslot(driver, extentTest, testName, testDataIteration, contactname, centre);
            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);

            //Navigating to activities and verifying language in payload

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["AsmLargeTourReques"].ToString());

            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            //generic.payloadverificationwithASM(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), "Regus", emailid, Ro["TourASMtest3"].ToString());
            generic.payloadverificationASMBusinessCentre(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), Ro["BussinessCentre"].ToString(), "Regus", emailid, Ro["TourASMtest3"].ToString());


            //CLOSE TOUR
            generic.NavigateTourfrompayload(driver, extentTest, testName, testDataIteration);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Action1"].ToString());

            login.Logout(driver, extentTest, testName, testDataIteration);

            //}
            //catch (Exception e)
            //{
            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");

            //    Type thisType = this.GetType();

            //    object testCall = this;

            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            //}
        }




        public static IEnumerable<object[]> RTA_16539()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }


        /// <summary>

        ///[CRM-2809], [CRM-7618]-Verify whether the communication request is generated for assigned user when the tour is scheduled.

        /// </summary>

        [TestCategory("Regression"), TestCategory("GlobalP11"), TestCategory("HF"), TestCategory("25Nov"), TestCategory("08Sep"), TestCategory("GlobalP11A"), TestCategory("25Sep2020")]

        [TestProperty("TestcaseID", "RTA-16539")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_16539), DynamicDataSourceType.Method)]

        public void RTA_16539_Verifycommunreqttourschd(DataRow Ro)
        {
            //try { 

            //{ 

            // Testdata:
            string newdate = DateTime.Today.ToString("MM/dd/yyyy").Replace("-", "/");
            string Date = DateTime.Today.ToString("yyyy-MM-dd");
            Console.WriteLine(Date);
            string now = System.DateTime.Now.ToString();

            // Login to the application
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            //Creating new Contact
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            string time = System.DateTime.Now.ToString();
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            string lastname = generic.CreateNewContactforPayloadVerification(driver, extentTest, testDataIteration, testName, Ro["Contactname"].ToString(), now, Ro["Email"].ToString(), "");
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");
            generic.SelectOpportunityheader(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString(), "Form:contact");
            generic.scrollDownOpportunity(driver, extentTest, testName, testDataIteration);
            generic.EnterContactNumber(driver, extentTest, testName, testDataIteration, "+919497852369");

            //Click Opportunity and create new Opp
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, "New");
            string people = "5";
            generic.CreateNewOpportunityEntetrprisewithprevcontact1(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["MinorSource"].ToString(), contactname, Ro["BussinessCentre"].ToString(), people);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ThinkTime(5);
            WaitUntil(driver, Control("OppReference", "EnterpriseSales"), 60);
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Navigating to activities and verifying owner in payload
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, "Activities");
            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["AsmTourReques"].ToString());
            //generic.HidingPureCloud(driver, extentTest, testName, testDataIteration);

            // Verify the owner            
            generic.CommrequestHeaderverification(driver, extentTest, testName, testDataIteration, Ro["Commreqowner"].ToString());

            //verifying scheduled call back request payload
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            // generic.payloadverificationwithASM(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), "Regus", Ro["Email"].ToString(), Ro["TourASMtest3"].ToString());
            generic.payloadverificationASMBusinessCentre1(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), Ro["BussinessCentre"].ToString(), "Regus", Ro["Email"].ToString(), Ro["TourASMtest3"].ToString());

            // Go back to Opportunity
            generic.Backtooppotunity(driver, extentTest, testName, testDataIteration);
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
            generic.OpenExistingTour(driver, extentTest, testName, testDataIteration);
            ThinkTime(3);

            // Reschedule the tour
            generic.RescheduleTour(driver, extentTest, testName, testDataIteration, "Form:pro_tour");
            generic.ResheduleBookaslot(driver, extentTest, testName, testDataIteration, "", "");

            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Navigating to activities and verifying language in payload
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, "Activities");
            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["AsmTourReques"].ToString());

            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverificationwithASM(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), "Regus", Ro["Email"].ToString(), Ro["TourASMtest3"].ToString());

            generic.Logout(driver, extentTest, testName, testDataIteration);

            // //Login as area Director
            generic.Loginafterlogout2(driver, extentTest, testName, testDataIteration, Ro["UserNameAD"].ToString(), Ro["PasswordAD"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            // Open tour from calendar
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, "Dashboards");
            //generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, Ro["Role"].ToString());
            generic.OpenCalanderFromTourManagement(driver, extentTest, testName, testDataIteration);
            generic.SelectingCityDairyFromBTCalander(driver, extentTest, testName, testDataIteration, Ro["Reassigncity"].ToString());
          
            // Dear and drop the tour to another ASM
            generic.DragAndDropToUnAvailableASM(driver, extentTest, testName, testDataIteration);

            driver.Close();

            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);

            //Navigating to activities and verifying language in payload
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["AsmLargeTourReques"].ToString());

            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverificationwithASM(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), "Regus", Ro["Email"].ToString(), Ro["TourASMtest3"].ToString());

            // Logout 
            generic.Logout(driver, extentTest, testName, testDataIteration);

            //}
            //catch (Exception e)
            //{
            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");
            //}

        }




        public static IEnumerable<object[]> RTA_16540()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        ///TEST[CRM-2798]Verify whether the communication request is generated for assigned user when the tour is rescheduled.
        /// CRM-2809

        /// </summary>

        [TestCategory("Regression"), TestCategory("GlobalP11Fail"), TestCategory("2211"), TestCategory("GlobalP11_2")]

        [TestProperty("TestcaseID", "RTA-16540")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_16540), DynamicDataSourceType.Method)]



        public void RTA_16540_Verifycommreqst(DataRow Ro)
        {
            //try

            //{              

            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);


            string emailid = "AutomatCompany" + Time + "@gmail.com";

            string now = System.DateTime.Now.ToString();


            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");
            //Creating new Opp
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);

            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString()); //start
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "13";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, Ro["BussinessCentre"].ToString(), People);

            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");


            ThinkTime(5);
            // Refresh the page
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());


            //Verifying Opportunity status and Language

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString()); //end


            //Booking a tour
            //generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            //generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, "OPP-437690"); 
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());

            string centre = Ro["BussinessCentre"].ToString();



            generic.Bookaslot(driver, extentTest, testName, testDataIteration, "", oppname); //change

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Re-schedule tour
            //generic.Backtooppotunity(driver, extentTest, testName, testDataIteration);
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
            generic.OpenExistingTour(driver, extentTest, testName, testDataIteration);
            generic.RescheduleTour(driver, extentTest, testName, testDataIteration, "");
            generic.ResheduleBookaslot(driver, extentTest, testName, testDataIteration, "", centre); //change

            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname); //change

            //Navigating to activities and verifying language in payload

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["AsmLargeTourReques"].ToString());

            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverificationwithASM(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), "Regus", emailid, Ro["TourASMtest3"].ToString());


            //CLOSE TOUR
            generic.NavigateTourfrompayload(driver, extentTest, testName, testDataIteration);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Action1"].ToString());



            login.Logout(driver, extentTest, testName, testDataIteration);

            //}



            //catch (Exception e)

            //{



            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");

            //    Type thisType = this.GetType();

            //    object testCall = this;

            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            //}









        }






        public static IEnumerable<object[]> RTA_16762()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        //TEST[CRM - 2810] - Verify whether the communication request is generated for assigned user when the tour is re - assigned


        //  TEST[CRM - 2810//.

        /// </summary>
        /// 

        [TestCategory("Regression"), TestCategory("GlobalP11_2"), TestCategory("0911"), TestCategory("GlobalP11A")]

        [TestProperty("TestcaseID", "RTA-16762")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_16762), DynamicDataSourceType.Method)]


        public void RTA_16762_AreaDirectorDragAndDropTourForSameTimeSlots(DataRow Ro)
        {
            //try { 

            //{            
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);


            string emailid = "AutomatCompany" + Time + "@gmail.com";

            string now = System.DateTime.Now.ToString();


            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");

            //Creating new Opp


            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "13";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, Ro["BussinessCentre"].ToString(), People);

            string oppname = generic.GetOppID(driver, extentTest, testName, testDataIteration);



            ThinkTime(2);
            // Refresh the page
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Verifying Opportunity status and Language

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());

            string centre = Ro["ReassignCentre"].ToString();

            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());


            generic.Logout(driver, extentTest, testName, testDataIteration);

            //Login as area direstor

            //generic.LoginafterlogoutOld(driver, extentTest, testName, testDataIteration, Ro["UserNameAD"].ToString(), Ro["PasswordAD"].ToString());
            generic.Loginafterlogout2(driver, extentTest, testName, testDataIteration, Ro["UserNameAD"].ToString(), Ro["PasswordAD"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, "Dashboards");
            //generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, Ro["Role"].ToString());
            generic.OpenCalanderFromTourManagement(driver, extentTest, testName, testDataIteration);
            generic.SelectingCityDairyFromBTCalander(driver, extentTest, testName, testDataIteration, Ro["Reassigncity"].ToString());
            //generic.SelectingCentreFromBTCalander(driver, extentTest, testName, testDataIteration, Ro["ReassignCentre"].ToString());

            generic.DragAndDropToUnAvailableASM(driver, extentTest, testName, testDataIteration);

            driver.Close();

            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);

            //Navigating to activities and verifying language in payload

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["AsmLargeTourReques"].ToString());

            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverificationASMinfo(driver, extentTest, testName, testDataIteration, Ro["ReassignedASM"].ToString());



            generic.Logout(driver, extentTest, testName, testDataIteration);
            //}


            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}
        }



        public static IEnumerable<object[]> RTA_16765()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }


        //TEST[CRM - 2809] - Verify whether the communication request is generated for assigned user when the tour is re - assigned


        //  TEST[CRM - 2809//.

        /// </summary>
        /// 

        [TestCategory("Regression"), TestCategory("GlobalP11Fail"), TestCategory("0911"), TestCategory("GlobalP11_2")]

        [TestProperty("TestcaseID", "RTA-16765")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_16765), DynamicDataSourceType.Method)]


        public void RTA_16765_AreaDirectorDragAndDropTourForSameTimeSlots(DataRow Ro)
        {

            //try

            //{              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
            // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
            // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
            // generic.Logout(driver, extentTest, testName, testDataIteration);
            // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@gmail.com";
            string now = System.DateTime.Now.ToString();

            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");

            //Creating new Opp
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "4";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, Ro["BussinessCentre"].ToString(), People);
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ThinkTime(3);

            ////Verifying Opportunity status and Language
            //generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            //generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.Logout(driver, extentTest, testName, testDataIteration);

            //Login as area director
            generic.LoginafterlogoutOld(driver, extentTest, testName, testDataIteration, Ro["UserNameAD"].ToString(), Ro["PasswordAD"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, "Dashboards");
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, Ro["Role"].ToString());
            generic.OpenCalanderFromTourManagement(driver, extentTest, testName, testDataIteration);
            generic.SelectingCityDairyFromBTCalander(driver, extentTest, testName, testDataIteration, Ro["Reassigncity"].ToString());
            generic.SelectingCentreFromBTCalander(driver, extentTest, testName, testDataIteration, Ro["ReAsgcentre"].ToString());

            generic.DragAndDropToUnAvailableASM(driver, extentTest, testName, testDataIteration);

            driver.Close();

            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);

            //Navigating to activities and verifying language in payload

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["AsmTourReques"].ToString());

            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverificationASMinfo(driver, extentTest, testName, testDataIteration, Ro["ReassignedASM"].ToString());

            //    generic.Logout(driver, extentTest, testName, testDataIteration);
            //}


            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}
        }





        public static IEnumerable<object[]> RTA_7704()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        /// </summary>
        /// 

        //[TEST] [CRM-3222] Verify that CRM will automatically link who owns the opportunity to the opportunity with a connection role of “Sales Assist” after booking a tour (Logged in user = Opp owner)

        [TestCategory("Regression"), TestCategory("GlobalP11_2"), TestCategory("BAT"), TestCategory("HF3"), TestCategory("PriorReg"), TestCategory("GlobalP11A"), TestCategory("HF")]

        [TestProperty("TestcaseID", "RTA-7704")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_7704), DynamicDataSourceType.Method)]


        public void RTA_7704_saleassist(DataRow Ro)
        {
            try

            {

                // Testdata:
                string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@gmail.com";
            string now = System.DateTime.Now.ToString();

            // Login to the application
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            ThinkTime(10);
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");

            //Creating new Opp
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "8";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, Ro["BussinessCentre"].ToString(), People);
            ThinkTime(3);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            string oppname = generic.GetOppID(driver, extentTest, testName, testDataIteration);
            ThinkTime(3);

            // Refresh the page
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre"].ToString();
                generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());           

            // Navigate to the connections
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["Connections"].ToString());

            ///verify opp owner connected to opportunity by sale assist roile
            generic.VerifysalesassistroleNew(driver, extentTest, testName, testDataIteration, Ro["User1"].ToString(), Ro["Rolesaleassist"].ToString());
           
            // Logout from the application
            login.Logout(driver, extentTest, testName, testDataIteration);



            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }


        }



        public static IEnumerable<object[]> RTA_7705()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        //TEST[CRM - 2809] -Verify that CRM will automatically link who owns the opportunity to the opportunity with a connection role of “Sales Assist” after booking a tour (Logged in user != Opp owner)


        //  TEST[CRM - 3222//.

        /// </summary>
        /// 

        [TestCategory("Regression"), TestCategory("GlobalP11Failed"), TestCategory("0911"), TestCategory("GlobalP11_2"), TestCategory("25Sep2020")]

        [TestProperty("TestcaseID", "RTA-7705")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_7705), DynamicDataSourceType.Method)]


        public void RTA_7705_OppConnectionRoleofSalesAssist(DataRow Ro)
        {
            //try

            //{              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
            // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
            // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
            // generic.Logout(driver, extentTest, testName, testDataIteration);
            // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@gmail.com";
            string now = System.DateTime.Now.ToString();


            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");


            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            //Creating new Opp
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "13";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, Ro["salesassistcentre"].ToString(), People);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ThinkTime(10);
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

            // Logout
            //generic.Logout(driver, extentTest, testName, testDataIteration);

            ////Login as area director
            ////generic.LoginafterlogoutOld(driver, extentTest, testName, testDataIteration, Ro["UserNameAD"].ToString(), Ro["PasswordAD"].ToString());
            //generic.Loginafterlogout2(driver, extentTest, testName, testDataIteration, Ro["UserName1"].ToString(), Ro["Password1"].ToString());
            //generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            //generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, "Dashboards");


            ////Navigating to Opportunity and selecting particular Opp record
            //generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            //generic.HidingPureCloud(driver, extentTest, testName, testDataIteration);
            //generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            //generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["salesassistcentre"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["Connections"].ToString());
            generic.selectConnectedTo(driver, extentTest, testName, testDataIteration, Ro["ConnectedTo"].ToString());
            //verify opp owner connected to opportunity by sale assist roile
            generic.Verifysalesassistrole(driver, extentTest, testName, testDataIteration, Ro["ConnectedTo"].ToString(), Ro["AsthisRole"].ToString());

            login.Logout(driver, extentTest, testName, testDataIteration);


            //}

            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}


        }

        public static IEnumerable<object[]> RTA_15392()
        {
            foreach (DataRow row in getTestCaseList("AreaDirector"))
            {
                yield return new object[] { row };
            }
        }

        //Verify the "Sales Assist" automation when a tour is booked an the Opportunity owner is from the Field Sales team

        //  TEST[CRM - 3222//.

        /// </summary>
        /// 

        [TestCategory("Regression"), TestCategory("GlobalP11_2"), TestCategory("2211"), TestCategory("GlobalP11A")]

        [TestProperty("TestcaseID", "RTA-15392")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_15392), DynamicDataSourceType.Method)]


        public void RTA_15392_Salesroles(DataRow Ro)
        {
            //try

            //{              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
            // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
            // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
            // generic.Logout(driver, extentTest, testName, testDataIteration);
            // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@gmail.com";
            string now = System.DateTime.Now.ToString();
            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["EntityContact"].ToString());
            generic.CreateNewContactDirector(driver, extentTest, testDataIteration, testName, Ro["Contactname"].ToString(), now);
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");


            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());

            ////Creating new Opp
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "8";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["MinorSource"].ToString(), contactname, Ro["BussinessCentre2"].ToString(), People);
            generic.ScriptErrorExcep(driver, extentTest, testName, testDataIteration);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            string oppname = generic.GetOppID(driver, extentTest, testName, testDataIteration);



            //Booking a tour

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());

            string centre = Ro["BussinessCentre2"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            //verify the owner is not connected with role salesassist
            //generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["Connections"].ToString());
            //generic.Verifyactivitynotpresent(driver, extentTest, testName, testDataIteration, Ro["SalesAssist"].ToString());

            //generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //verify tour owner
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
            generic.OpenExistingTour(driver, extentTest, testName, testDataIteration);
            ThinkTime(3);
            string owner = Element(driver, Control("tourowner", "Generic")).GetAttribute("title");
            generic.Verifytourowner(driver, extentTest, testName, testDataIteration, owner);
            //verify new opp owner
            generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);
            generic.VerifyAssignedUserActivity(driver, extentTest, testName, testDataIteration, owner);

            login.Logout(driver, extentTest, testName, testDataIteration);

            //}

            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}




        }




        public static IEnumerable<object[]> RTA_16279()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        /// <summary>
        /// Verify that Dynamics sends the preferred language of the Opportunity Contact to PureCloud when the Dialer - UnAttended.
        /// CRM-3366
        /// </summary>
        [TestCategory("08Sep"), TestCategory("Regression"), TestCategory("BAT"), TestCategory("GlobalP11_2"), TestCategory("CRM-3366"), TestCategory("GlobalP11A"), TestCategory("25Sep2020")]
        [TestProperty("TestcaseID", "RTA-16279")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_16279), DynamicDataSourceType.Method)]

        public void RTA_16279_VerifylangPureCloud(DataRow Ro)
        {
            //try
            //{              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
            // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
            // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
            // generic.Logout(driver, extentTest, testName, testDataIteration);
            // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            //Create new Opportunity for Book a tour
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);

            string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, "Oslo, Spaces Nydalen");
            //generic.LogoutPurecloud(driver, extentTest, testName, testDataIteration);

            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");


            //Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.VerifyStatusReason(driver, extentTest, testName, testDataIteration, "In Progress");
            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            //string centre = "Aalst, Erembodegem";
            string centre = "Oslo, Spaces Nydalen";
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Selecting Activities and filling values for tour outcome
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.VerifyTourCommRequestActivity(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle1"].ToString());
            generic.FillTourOutcome(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle1"].ToString(), Ro["SelectValue"].ToString(), Ro["TourOutcome"].ToString());

            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);

            //Navigating to activities and verifying language in payload
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.VerifyTourCommRequestActivity(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle2"].ToString());
            generic.VerificationinPayload(driver, extentTest, testName, testDataIteration, "UK English");


            login.Logout(driver, extentTest, testName, testDataIteration);
            //}

            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}

        }




        public static IEnumerable<object[]> RTA_16280()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        /// <summary>
        /// Verify that Dynamics sends the preferred language of the Opportunity Contact to PureCloud when the dialer has attended and contact is modfied.
        /// CRM-3366
        /// </summary>
        [TestCategory("08Sep"), TestCategory("28Aug"), TestCategory("Regression"), TestCategory("BAT"), TestCategory("GlobalP11_2"), TestCategory("CRM-3366"), TestCategory("GlobalP11A"), TestCategory("28Sep2020")]
        [TestProperty("TestcaseID", "RTA-16280")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_16280), DynamicDataSourceType.Method)]

        public void RTA_16280_PureCloudverify(DataRow Ro)
        {

            //try
            //{             

            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            //Create new Opportunity for Book a tour
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            string contactname = generic.CreateNewOppWithIsCustomerAccount(driver, extentTest, testName, testDataIteration, "Alberta, Calgary - Crowfoot Centre");
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

            //Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.VerifyStatusReason(driver, extentTest, testName, testDataIteration, "In Progress");

            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre2"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Selecting Activities and filling values for tour outcome
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.VerifyTourCommRequestActivity(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle1"].ToString());
            generic.FillTourOutcome(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle1"].ToString(), Ro["SelectValue"].ToString(), Ro["TourOutcome"].ToString());

            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);

            //Navigating to activities and verifying language in payload
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.VerifyTourCommRequestActivity(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle2"].ToString());
            generic.VerificationinPayload(driver, extentTest, testName, testDataIteration, "UK English");


            //Changing the language of the contact 
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "Active Contacts");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, contactname);
            //generic.SelectActiveCellContacts(driver, extentTest, testName, testDataIteration);
            generic.LanguageChangeinContact(driver, extentTest, testName, testDataIteration, "Norwegian");

            //Navigating to opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);
            //generic.SelectActiveCellContacts(driver, extentTest, testName, testDataIteration);

            //Booking another tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre1 = Ro["BussinessCentre2"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Selecting Activities and filling values for tour outcome
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.VerifynextTourCommRequestActivity(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle1"].ToString());
            generic.FillTourOutcome(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle1"].ToString(), Ro["SelectValue1"].ToString(), Ro["TourOutcome"].ToString());

            //Navigating to Opportunity and selecting particular Opp record
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Opportunities");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, oppname);


            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.VerifyPostTourCommRequestActivity(driver, extentTest, testName, testDataIteration, Ro["P1posttournattendeddialer"].ToString());
            generic.VerificationinPayload(driver, extentTest, testName, testDataIteration, Ro["Norwegian"].ToString());

            login.Logout(driver, extentTest, testName, testDataIteration);
            //}

            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}

        }


        public static IEnumerable<object[]> RTA_16289()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }


        /// <summary>
        ///[CRM-3022], [CRM-7612]-Verify whether the communication request for tour confirmation is generated for schedule and re-schedule tour in UK English language.
        /// </summary>

        [TestCategory("Regression"), TestCategory("GlobalP11_2"), TestCategory("HF"), TestCategory("HF3"), TestCategory("GlobalP11A"), TestCategory("28Sep2020")]
        [TestProperty("TestcaseID", "RTA-16289")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_16289), DynamicDataSourceType.Method)]
        public void RTA_16289_commrequestfortouruk(DataRow Ro)
        {
            try
            {
                // TestData
                string Date = DateTime.Today.ToString("dd-MM-yyyy");
                DateTime Now = DateTime.Now;
                string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
                string emailid = "AutomatCompany" + Time + "@gmail.com";
                string centre = Ro["BussinessCentre1"].ToString();

                // Login to the applicartion
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                // Navigate to Sales drop down
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);

                // Create new Contact
                string contactname = generic.CreateNewcontactOpportunity(driver, extentTest, testName, testDataIteration, Ro["BussinessCentre2"].ToString(), Ro["Language1"].ToString(), emailid);
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
               
                // Get the contactname
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
                                
                //Schedule a tour
                generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

                // Go to Tours and open the request
                generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());
                generic.searchrequest(driver, extentTest, testName, testDataIteration, Ro["P1searchtourrequest"].ToString());              
                generic.SelectingActiverequest(driver, extentTest, testName, testDataIteration);

                // Verify payload verification
                generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
                generic.payloadverificationforlang(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString());

                //Re-schedule tour
                generic.Backtooppotunity(driver, extentTest, testName, testDataIteration);
                generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
                generic.OpenExistingTour(driver, extentTest, testName, testDataIteration);
                generic.RescheduleTour(driver, extentTest, testName, testDataIteration, "");
                generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);

                // Open the new Comm request for the rescheduled tour
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
                //generic.NavigatetooppfromAppointment(driver, extentTest, testName, testDataIteration);
                generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());
                generic.searchrequest(driver, extentTest, testName, testDataIteration, Ro["P1searchtourrequest"].ToString());
                generic.reschedulecommrequest(driver, extentTest, testName, testDataIteration);
                //generic.tourcommrequestverify(driver, extentTest, testName, testDataIteration);
                generic.SelectingActiverequest(driver, extentTest, testName, testDataIteration);

                // Navigate to Payload and verify the details
                generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
                generic.payloadverification(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), centre, emailid, "");
                //generic.payloadverificationforlang(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString());

                // Logout from the application
                login.Logout(driver, extentTest, testName, testDataIteration);
            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }
        }



        public static IEnumerable<object[]> RTA_16292()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// [CRM-3022]-Verify whether the communication request for tour confirmation is generated for schedule and re-schedule tour in Norwegian language.
        /// </summary>

        [TestCategory("Regression"), TestCategory("GlobalP11"), TestCategory("0911"), TestCategory("GlobalP11A"), TestCategory("CreateOppEnt")]
        [TestProperty("TestcaseID", "RTA-16292")]

        [DataTestMethod]
        [DynamicData(nameof(RTA_16292), DynamicDataSourceType.Method)]
        public void RTA_16292_commrequestfortournorwegian(DataRow Ro)
        {
            //try
            //{

                string Date = DateTime.Today.ToString("dd-MM-yyyy");
                DateTime Now = DateTime.Now;
                string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
                string emailid = "AutomatCompany" + Time + "@gmail.com";
                string now = System.DateTime.Now.ToString();              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                                                                          // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                                                                          // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                                                                          // generic.Logout(driver, extentTest, testName, testDataIteration);
                                                                          // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
                string contactname = generic.CreateNewOpportunityEntetrprise(driver, extentTest, testDataIteration, testName, Ro["TestOpportunityAccountName"].ToString(), now, Ro["Country"].ToString(), Ro["TestCurrency"].ToString(), "Centre: Walk-in", Ro["TestMajorSource"].ToString(), "Acquisition", Ro["TestNewContact"].ToString());
                string oppname = generic.GetOppID(driver, extentTest, testName, testDataIteration);

                //Verifying Opportunity status and Language
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
                generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

                //Booking a tour
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());

                generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());


                generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());
                generic.searchrequest(driver, extentTest, testName, testDataIteration, Ro["P1searchtourrequest"].ToString());
                generic.SelectingActiverequest(driver, extentTest, testName, testDataIteration);

                generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
                generic.payloadverification(driver, extentTest, testName, testDataIteration, Ro["Languagecode"].ToString(), "Oslo, Spaces Nydalen", emailid, "");
                //Re-schedule tour
                generic.Backtooppotunity(driver, extentTest, testName, testDataIteration);
                generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
                generic.OpenExistingTour(driver, extentTest, testName, testDataIteration);
                generic.RescheduleTour(driver, extentTest, testName, testDataIteration, "");
                //generic.ResheduleBookaslot(driver, extentTest, testName, testDataIteration, contactname, "Oslo, Spaces Nydalen");
                generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                //generic.NavigatetooppfromAppointment(driver, extentTest, testName, testDataIteration);
                generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());
                generic.searchrequest(driver, extentTest, testName, testDataIteration, Ro["P1searchtourrequest"].ToString());
                generic.reschedulecommrequest(driver, extentTest, testName, testDataIteration);
                generic.tourcommrequestverify(driver, extentTest, testName, testDataIteration);
                generic.SelectingActiverequest(driver, extentTest, testName, testDataIteration);
                generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
                generic.payloadverification(driver, extentTest, testName, testDataIteration, Ro["P1languagecodenor"].ToString(), "Oslo, Spaces Nydalen", emailid, "");
                login.Logout(driver, extentTest, testName, testDataIteration);
            //}

            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}
        }




        public static IEnumerable<object[]> RTA_16719()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>

        ///[CRM-3349]-Verify whether the dialer request is generated when the tour is schedule in dynamics.
        /// </summary>

        [TestCategory("Regression"), TestCategory("GlobalP11"), TestCategory("25Nov1"), TestCategory("28Aug"), TestCategory("GlobalP11A"), TestCategory("2810"), TestCategory("2910")]

        [TestProperty("TestcaseID", "RTA-16719")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_16719), DynamicDataSourceType.Method)]
        public void RTA_16719_Verifypostcommrequesttourschedule(DataRow Ro)
        {
            //try

            //{
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                string Date = DateTime.Today.ToString("dd-MM-yyyy");
                DateTime Now = DateTime.Now;
                string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);


                string emailid = "AutomatCompany" + Time + "@gmail.com";

                string now = System.DateTime.Now.ToString();
                string PT = "POST-TOUR";
                //Select Contacts entity and create new
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
                string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");



                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                //Creating new Opp
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
                generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
                string People = "10";
                generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, Ro["BussinessCentre"].ToString(), People);

                string oppname = generic.GetOppID(driver, extentTest, testName, testDataIteration);


                ThinkTime(2);
                WaitUntil(driver, Control("FieldVerification2", "Refresh", "Opportunity"), 120);

                // Refresh the page
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

                //Verifying Opportunity status and Language

                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

                generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

                //Booking a tour

                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());

                string centre = Ro["BussinessCentre"].ToString();
                //string contactname = "";
                generic.Bookaslot(driver, extentTest, testName, testDataIteration, oppname, oppname);

                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());



                generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());
                generic.searchrequest(driver, extentTest, testName, testDataIteration, Ro["P1posttournattendeddialer"].ToString());
                // generic.Verifyactivitynotpresent(driver, extentTest, testName, testDataIteration, Ro["P1posttournattendeddialer"].ToString());
                //Tour record for yes

                generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
                generic.TourOutcomesave(driver, extentTest, testName, testDataIteration, Ro["CreatedTour"].ToString(), Ro["P1show"].ToString());
                generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);
                //Verify post tour yes activity
                generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());
                generic.searchrequest(driver, extentTest, testName, testDataIteration, Ro["P1posttournattendeddialer"].ToString());
                generic.dialercommrequest(driver, extentTest, testName, testDataIteration);
                generic.posttourcommrequestverify(driver, extentTest, testName, testDataIteration, Ro["P1posttournattendeddialer"].ToString());
                //Tour record for NO
                generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
                generic.TourOutcome(driver, extentTest, testName, testDataIteration, Ro["CreatedTour"].ToString(), Ro["P1noshow"].ToString(), Ro["P1noshowreason"].ToString());
                generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());

                //Verify post tour No activity


                generic.searchrequest(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle2"].ToString());
                generic.dialercommrequest(driver, extentTest, testName, testDataIteration);
                generic.posttourcommrequestverify(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle2"].ToString());
                generic.SelectingActiverequest(driver, extentTest, testName, testDataIteration);
                generic.verifydialeraction(driver, extentTest, testName, testDataIteration, Ro["P1Action"].ToString());

                //Verifying Payload fields
                generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
                generic.payloadverification01(driver, extentTest, testName, testDataIteration, Ro["Language1"].ToString(), Ro["regusbrand"].ToString(), PT, "");
                login.Logout(driver, extentTest, testName, testDataIteration);

            //}
            //catch (Exception e)
            //{
            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");
                

            //}
        }






        public static IEnumerable<object[]> RTA_16806()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        /// <summary>

        ///[CRM-3365] Verify that the Dynamics sends the email address or team name of the Owner of the "In Progress" Opportunity to PureCloud if the Owner gets updated (via assigning)

        /// </summary>

        [TestCategory("Regression"), TestCategory("GlobalP11"), TestCategory("BAT"), TestCategory("0911")]

        [TestProperty("TestcaseID", "RTA-16806")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_16806), DynamicDataSourceType.Method)]

        public void RTA_16806_CommRqstPTattendedopp(DataRow Ro)
        {



            //try
            //{
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                string Date = DateTime.Today.ToString("dd-MM-yyyy");
                DateTime Now = DateTime.Now;
                string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
                string emailid = "AutomatCompany" + Time + "@mailinator.com";

                //Create new Opportunity for Book a tour
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
                string contactname = generic.CreateNewcontactOpportunity(driver, extentTest, testName, testDataIteration, Ro["BussinessCentre"].ToString(), Ro["Language1"].ToString(), emailid);
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.AssignRecordAnotherUser(driver, extentTest, testName, testDataIteration, Ro["AssignTo"].ToString(), Ro["User1"].ToString());

                //Verifying Opportunity status and Language
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
                generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

                //Booking a tour
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
                string centre = Ro["BussinessCentre"].ToString();
                generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Selecting Activities and filling values for tour outcome
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
            generic.TourOutcome(driver, extentTest, testName, testDataIteration, Ro["CreatedTour"].ToString(), Ro["P1noshow"].ToString(), Ro["P1noshowreason"].ToString());
            //generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());

            //Navigating to activities and verifying language in payload
            generic.searchrequest(driver, extentTest, testName, testDataIteration, Ro["P1posttournoshow"].ToString());
            generic.posttourcommrequestverify(driver, extentTest, testName, testDataIteration, Ro["P1posttournoshow"].ToString());
            generic.SelectingActiverequest(driver, extentTest, testName, testDataIteration);

            //Verifying Payload fields
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverification01(driver, extentTest, testName, testDataIteration, Ro["User1"].ToString(), "", "", "");

            login.Logout(driver, extentTest, testName, testDataIteration);
            //}
            //catch (Exception e)

            //{



            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");

            //    Type thisType = this.GetType();

            //    object testCall = this;

            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            //}
        }




        public static IEnumerable<object[]> RTA_17643()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        ///TEST[CRM-2808],[CRM-7818]-Verify that "ASM, AD Internal Handover Email - Call back Large deal X workstations" Comm Request is generated when a new large deal Opportunity is created
        /// </summary>

        [TestCategory("HF"), TestCategory("Regression"), TestCategory("BAT"), TestCategory("GlobalP1_S2"), TestCategory("vg"), TestCategory("11Sep")]

        [TestProperty("TestcaseID", "RTA-17643")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_17643), DynamicDataSourceType.Method)]

        public void RTA_17643_largedealcallback(DataRow Ro)
        {
            //try

            //{

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@mailinator.com";
            string now = System.DateTime.Now.ToString();              
            
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");


            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            //Creating new Opp
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "13";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, Ro["BussinessCentre"].ToString(), People);

            // Refresh the page
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            ThinkTime(3);
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre = Ro["BussinessCentre"].ToString();
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Navigating to activities and verifying language in payload
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["ActivityTitleDialer"].ToString());

            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            // generic.payloadverificationwithASM(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), "Regus", emailid, Ro["TourASMtest3"].ToString());
            generic.payloadverificationASMBusinessCentre(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), Ro["BussinessCentre"].ToString(), "Regus", emailid, Ro["TourASMtest3"].ToString());

            ////CLOSE TOUR
            //generic.NavigateTourfrompayload(driver, extentTest, testName, testDataIteration);
            //generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Mark Complete"].ToString());

            login.Logout(driver, extentTest, testName, testDataIteration);

            //}

            //catch (Exception e)

            //{
            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");
            //}

        }




        public static IEnumerable<object[]> RTA_8931()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// CRM-3487

        /// </summary>

        [TestCategory("JMS"), TestCategory("Regression"), TestCategory("1410"), TestCategory("GlobalP1_S2"), TestCategory("09Sep")]

        [TestProperty("TestcaseID", "RTA-8931")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_8931), DynamicDataSourceType.Method)]
        public void RTA_8931_Verifycrntremiageaforregusbrand(DataRow Ro)
        {
            //try

            //{              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
            // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
            // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
            // generic.Logout(driver, extentTest, testName, testDataIteration);
            // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);

            string emailid = "AutomatCompany" + Time + "@gmail.com";
            string centre = "Stuttgart, Konigstrasse";
            string now = System.DateTime.Now.ToString();


            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            // Select Contacts entity and create new
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");

            //Creating new Opp


            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "6";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, centre, People);

            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

            // Refresh the page
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Verifying Opportunity status and Language

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);

            ThinkTime(5);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["emailconnRqst"].ToString());
            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverification(driver, extentTest, testName, testDataIteration, "", "", "", "", "Centre.CentreImage", "http://assets.regus.com/images/email/Regus.jpg", "Centre.BrandImage", "https://assets.regus.com/images/");
            //CLOSE TOUR
            generic.NavigateTourfrompayload(driver, extentTest, testName, testDataIteration);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Action1"].ToString());
            login.Logout(driver, extentTest, testName, testDataIteration);
            //}
            //catch (Exception e)
            //{
            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");

            //    Type thisType = this.GetType();

            //    object testCall = this;

            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            //}
        }





        /// CRM-4528
        public static IEnumerable<object[]> RTA_16960()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// </summary>

        [TestCategory("0409"), TestCategory("JMS"), TestCategory("BAT"), TestCategory("Regression"), TestCategory("25Nov1"), TestCategory("GlobalP1_S2"), TestCategory("11Sep")]

        [TestProperty("TestcaseID", "RTA-16960")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_16960), DynamicDataSourceType.Method)]
        public void RTA_16960_Verifycentreimageandbrandimageinpayload(DataRow Ro)
        {
            //try

            //{              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
            // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
            // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
            // generic.Logout(driver, extentTest, testName, testDataIteration);
            // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);

            string emailid = "AutomatCompany" + Time + "@mailinator.com";
            string centre = "Alberta, Calgary - Crowfoot Centre";
            string now = System.DateTime.Now.ToString();

            // Select Contacts entity and create new

            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");

            //Creating new Opp


            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "6";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, centre, People);

            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

            // Refresh the page
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Verifying Opportunity status and Language

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);

            ThinkTime(5);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["emailconnRqst"].ToString());
            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverification(driver, extentTest, testName, testDataIteration, "", "", "", "", "Centre.CentreImage", "http://assets.regus.com/images/email/Regus.jpg", "Centre.BrandImage", "https://assets.regus.com/images/");

            login.Logout(driver, extentTest, testName, testDataIteration);
            //}
            //catch (Exception e)
            //{
            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");

            //    Type thisType = this.GetType();

            //    object testCall = this;

            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            //}
        }




        /// CRM-3457
        public static IEnumerable<object[]> RTA_12077()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// </summary>

        [TestCategory("JMS"), TestCategory("GlobalP1_S2"), TestCategory("BAT"), TestCategory("27Aug"), TestCategory("Regression")]

        [TestProperty("TestcaseID", "RTA-12077")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_12077), DynamicDataSourceType.Method)]
        public void RTA_12077_VerifynocallprocessedifphonesettoNo(DataRow Ro)
        {
            //try

            //{              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
            // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
            // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
            // generic.Logout(driver, extentTest, testName, testDataIteration);
            // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);

            string emailid = "AutomatCompany" + Time + "@gmail.com";
            string centre = "Stuttgart, Konigstrasse";
            string now = System.DateTime.Now.ToString();


            // Select Contacts entity and create new

            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.Disablecontactdetails(driver, extentTest, testName, testDataIteration, Ro["ContactInfo"].ToString(), Ro["ContactDetails"].ToString(), "phone");

            //Creating new Opp

            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "6";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, centre, People);

            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());

            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);

            // Refresh the page
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Navigating to activities and verifying language in payload
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, "No answer to Post Tour Call");
            //Verifying Payload response Request not processed
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverifyrqstnotprocessed(driver, extentTest, testName, testDataIteration, Ro["unabletoprcessRqst"].ToString(), "No answer to Post Tour Call");
            //CLOSE TOUR 
            generic.NavigateTourfrompayload(driver, extentTest, testName, testDataIteration);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Action1"].ToString());
            login.Logout(driver, extentTest, testName, testDataIteration);





            //}
            //catch (Exception e)
            //{
            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");

            //    Type thisType = this.GetType();

            //    object testCall = this;                                                                                                                                               

            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            //}
        }

        public static IEnumerable<object[]> RTA_21881()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        // <summary>
        ///To Verify whether the user cant close the opportunity as Won,when payment received is No.
        //////CRM 3372
        /// </summary>
        [TestCategory("smoke"), TestCategory("BAT"), TestCategory("HF"), TestCategory("HF3"),  TestCategory("GlobalP1_S2"), TestCategory("Regression")]
        [TestProperty("TestcaseID", "RTA-21881")]
        [DataTestMethod]
        [DynamicData(nameof(RTA_21881), DynamicDataSourceType.Method)]
        public void RTA_21881_TourSecondayOpp(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());

            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            string now = System.DateTime.Now.ToString();
            string emailid = "crm.test2@regus.com";

            // Create Account
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity2"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.SelectNewTab(driver, extentTest, testName, testDataIteration);
            string Brokeraccount = generic.CreateNewAccountWithCustomerTypeNew(driver, extentTest, testDataIteration, testName, Ro["TestAccountLink"].ToString(), now, Ro["CustomerType2"].ToString());
            generic.Movetobroker(driver, extentTest, testName, testDataIteration, Ro["CustomerType2"].ToString());

            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            //  string contactname1 = generic.CreateNewContactEnterpriseSalesForBroker(driver, extentTest, testDataIteration, testName, Ro["Contactname"].ToString(), now, "crm.test4@regus.com", Brokeraccount);
            string contactname = generic.CreateNewContactEnterpriseSalesForBroker(driver, extentTest, testDataIteration, testName, Ro["Contactname"].ToString(), now, emailid, Brokeraccount);

            // generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);

            //First Lead
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity3"].ToString());
            string firstleadname = generic.DuplicateLead(driver, extentTest, testName, testDataIteration, "Lead" + now, emailid, Ro["Product"].ToString(), "New", Ro["Source"].ToString(), Ro["MajorSource"].ToString(), Ro["MinorSource"].ToString(), "Oslo, Spaces Nydalen", contactname);

            string Time1 = System.DateTime.Now.ToString();
            //Second Lead
            string secondleadname = generic.DuplicateLead1(driver, extentTest, testName, testDataIteration, "Lead" + Time1, emailid, Ro["Product"].ToString(), "New1", firstleadname, Ro["Source"].ToString(), Ro["MajorSource"].ToString(), Ro["MinorSource"].ToString(), "Oslo, Spaces Nydalen", contactname);

            ////navigate to secondary opp.
            //generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, "Secondary Opportunity");
            //generic.OpensecondayOpp(driver, extentTest, testName, testDataIteration);


            // Qualify first lead
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity3"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Leads");
            generic.RecordGlobalSearch(driver, extentTest, testName, testDataIteration, firstleadname);
            generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
            //generic.selectsecondrecord(driver, extentTest, testName, testDataIteration);
            generic.NavigateToRelatedTabEntitiesLead(driver, extentTest, testName, testDataIteration, Ro["RelatedTabEntity"].ToString());


            ThinkTime(5);
            generic.OppSelect(driver, extentTest, testName, testDataIteration);
            string oppreference1 = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, "Secondary Opportunity");
            generic.verifyoriginatingleadinOpp(driver, extentTest, testName, testDataIteration, "New1 Lead26-11-2020 10:32:39");
            ThinkTime(4);
            string oppreference = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");

           
            // Link Broker Account and contact
            //  generic.LinkBrokerdetailsNew(driver, extentTest, testName, testDataIteration, Brokeraccount, contactname);


            // generic.verifytimelineemail(driver, extentTest, testName, testDataIteration);

            //Log out
            generic.Logout(driver, extentTest, testName, testDataIteration);
                //generic.Loginafterlogout(driver, extentTest, testName, testDataIteration, emailid, Ro["PasswordAD"].ToString());
                //generic.NavigateToOutlook(driver, extentTest, testName, testDataIteration);
                //generic.SearchAndVerifysystemEmailOutlook(driver, extentTest, testName, testDataIteration, contactname, contactname, oppreference, "Thank you for sending us your referral", "Regus Broker Team", oppreference1);


                ////Logout
                //login.Logout(driver, extentTest, testName, testDataIteration);


            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");
            }

        }




        public static IEnumerable<object[]> RTA_18010()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        // <summary>
        ///To Verify whether the user cant close the opportunity as Won,when payment received is No.
        //////CRM 3372
        /// </summary>
        [TestCategory("28Sep2020"), TestCategory("25Nov"), TestCategory("smoke"), TestCategory("RerunMay9"), TestCategory("BAT"), TestCategory("GlobalP1_S2"), TestCategory("Regression")]
        [TestProperty("TestcaseID", "RTA-18010")]
        [DataTestMethod]
        [DynamicData(nameof(RTA_18010), DynamicDataSourceType.Method)]
        public void RTA_18010_TourSecondayOpp(DataRow Ro)
        {
            //try
            //{
            string emailid = "crm.test2@regus.com";              //  login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                                                                 // generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                                                                 // generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                                                                 // generic.Logout(driver, extentTest, testName, testDataIteration);
                                                                 // login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            string Time = System.DateTime.Now.ToString();

            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.CreateNewContactDirectSalesUser(driver, extentTest, testDataIteration, testName, Ro["Contactname"].ToString(), Time);
            generic.scrollDownContactPage(driver, extentTest, testName, testDataIteration);
            generic.scrollDownContactPage(driver, extentTest, testName, testDataIteration);
            generic.Movetobroker(driver, extentTest, testName, testDataIteration, Ro["CustomerType2"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");

            //First Lead
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity3"].ToString());
            string firstleadname = generic.DuplicateLead(driver, extentTest, testName, testDataIteration, "Lead" + Time, Ro["Email"].ToString(), Ro["Product"].ToString(), "New", Ro["Source"].ToString(), Ro["MajorSource"].ToString(), Ro["MinorSource"].ToString(), "Oslo, Spaces Nydalen", contactname);
            string Time1 = System.DateTime.Now.ToString();
            //Second Lead
            string secondleadname = generic.DuplicateLead1(driver, extentTest, testName, testDataIteration, "Lead" + Time1, Ro["Email"].ToString(), Ro["Product"].ToString(), "New1", firstleadname, Ro["Source"].ToString(), Ro["MajorSource"].ToString(), Ro["MinorSource"].ToString(), "Oslo, Spaces Nydalen", contactname);


            // Qualify first lead
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity3"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "All Leads");
            generic.OpenRecordGlobalSearch(driver, extentTest, testName, testDataIteration, firstleadname);
            generic.NavigateToRelatedTabEntitiesLead(driver, extentTest, testName, testDataIteration, Ro["RelatedTabEntity"].ToString());
            generic.OppSelect(driver, extentTest, testName, testDataIteration);

            //Verification of first lead
            generic.PrimaryOppVerify(driver, extentTest, testName, testDataIteration, "Oslo, Spaces Nydalen");
            generic.OppSelect(driver, extentTest, testName, testDataIteration);


            string People = "13";
            generic.AddingNoofPeopleinOpp(driver, extentTest, testName, testDataIteration, People);
            string oppname = generic.GetOppID(driver, extentTest, testName, testDataIteration);

            // Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, "TestContact", oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, "Refresh");

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, "Activities");

            generic.VerifyTourActivitiesSecondaryopp(driver, extentTest, testName, testDataIteration);

            //Selecting Activities and filling values for tour outcome
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());


            generic.TourOutcomeArrive(driver, extentTest, testName, testDataIteration, Ro["CreatedTour"].ToString(), Ro["P1show"].ToString());
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            generic.scrollDownOpportunity(driver, extentTest, testName, testDataIteration);

            generic.VerifyPrimaryOpp(driver, extentTest, testName, testDataIteration);

            // generic.VerifyActivitiesPrimaryOpp(driver, extentTest, testName, testDataIteration);

            //navigate to secondary opp.
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, "Secondary Opportunity");
            generic.OpensecondayOpp(driver, extentTest, testName, testDataIteration);

            //Verification of secondary opps cancelled comm request.
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, "Activities");
            generic.VerifySecOppCommRequstCancel(driver, extentTest, testName, testDataIteration);

            ////Log out
            //generic.Logout(driver, extentTest, testName, testDataIteration);
            //generic.Loginafterlogout(driver, extentTest, testName, testDataIteration, emailid, Ro["Password1"].ToString());
            //generic.NavigateToOutlook(driver, extentTest, testName, testDataIteration);
            //generic.SearchAndVerifysystemEmailOutlook(driver, extentTest, testName, testDataIteration, contactname, contactname, oppreference, "Thank you for sending us your referral", "Regus Broker Team");

            // Log out
            generic.Logout(driver, extentTest, testName, testDataIteration);


            //}

            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");
            //}

        }




        public static IEnumerable<object[]> RTA_5079()
        {
            foreach (DataRow row in getTestCaseList("AreaDirector"))
            {
                yield return new object[] { row };
            }
        }

        /// <summary>
        /// Test Method to check the calendar gets grey to indicate the date and time in past is not available to book now.
        /// </summary>
        [TestCategory("BookaTour"), TestCategory("Sales Agent"), TestCategory("BAT"), TestCategory("25Nov1"), TestCategory("25AugFixed"), TestCategory("RerunMay120"), TestCategory("SprintUN")]
        [TestProperty("TestcaseID", "RTA-5079")]
        [DataTestMethod]
        [DynamicData(nameof(RTA_5079), DynamicDataSourceType.Method)]
        public void RTA_5079_VerifySameColourBookedASM(DataRow Ro)
        {
            //try
            //{              
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown1(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            string parent = driver.CurrentWindowHandle;
            //generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["BussinessCentre2"].ToString());
            string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
            generic.BookTourBtn(driver, extentTest, testName, testDataIteration, "Book Tour");
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, "", oppname);
            //generic.VerifyColorPastDate(driver, extentTest, testDataIteration, testName);
            generic.VerifyBookTourSlotColurDifferent(driver, extentTest, testDataIteration, testName);
            generic.CancelTour(driver, extentTest, testDataIteration, testName);
            driver.SwitchTo().Window(parent);
            //generic.VerifyColourBookedAM(driver, extentTest, testName, testDataIteration, "rgb(75, 194, 80)");
            login.Logout(driver, extentTest, testName, testDataIteration);

            //}
            //catch (Exception e)
            //{

            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
            //    Type thisType = this.GetType();
            //    object testCall = this;
            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            //}
        }


        public static IEnumerable<object[]> RTA_5080()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }

        /// <summary>
        /// Test Method to check the calendar gets grey to indicate the date and time in past is not available to book now.
        /// </summary>
        [TestCategory("smoke"), TestCategory("BookaTour"), TestCategory("Validation1"), TestCategory("BAT"), TestCategory("RerunMay9"), TestCategory("25AugFixed"), TestCategory("Rerun-24062020"), TestCategory("SprintUN")]
        [TestProperty("TestcaseID", "RTA-5080")]
        [DataTestMethod]
        [DynamicData(nameof(RTA_5080), DynamicDataSourceType.Method)]
        public void RTA_5080_VerifyOrangreColourBookedBySA(DataRow Ro)
        {
            try
            {
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserNametour"].ToString(), Ro["Passwordtour"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sales");
                generic.DeleteTourFn(driver, extentTest, testName, testDataIteration);
                generic.Logout(driver, extentTest, testName, testDataIteration);
                login.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());
                //login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown1(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
                string parent = driver.CurrentWindowHandle;
                generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
                string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration, Ro["Centre"].ToString());
                string oppname = Element(driver, Control("OppReference", "EnterpriseSales")).GetAttribute("title");
                generic.BookTourBtn(driver, extentTest, testName, testDataIteration, "Book Tour");
                try
                {
                    string centre = Ro["Centre"].ToString();
                    generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
                    generic.VerifyBookTourSlotColurDifferent(driver, extentTest, testDataIteration, testName);
                    generic.CancelTour(driver, extentTest, testName, testDataIteration);
                    driver.SwitchTo().Window(parent);
                    login.Logout(driver, extentTest, testName, testDataIteration);
                }
                catch (Exception e)
                {

                    AddLog(driver, extentTest, testName, testDataIteration, "Info", "Exception handled" + e.Message, "Exception handled");
                }



            }
            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");

            }
        }





        public static IEnumerable<object[]> RTA_16514()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }


        ///[CRM-2810]-Verify whether the communication request is generated for assigned user when the large deal tour is scheduled.
        /// </summary>

        [TestCategory("Regression"), TestCategory("25Nov"), TestCategory("BAT"), TestCategory("08Sep"), TestCategory("GlobalP11"), TestCategory("GlobalP11A"), TestCategory("24Nov")]

        [TestProperty("TestcaseID", "RTA-16514")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_16514), DynamicDataSourceType.Method)]

        public void RTA_16514_Verifycommuequestgeneratedtourscheduled(DataRow Ro)
        {
            //try

            //{             
            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);

            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            string emailid = "AutomatCompany" + Time + "@gmail.com";

            string now = System.DateTime.Now.ToString();


            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");

            //Creating new Opp


            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "13";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, Ro["BussinessCentre"].ToString(), People);

            string oppname = generic.GetOppID(driver, extentTest, testName, testDataIteration);



            ThinkTime(20);
            // Refresh the page
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());


            //Verifying Opportunity status and Language

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());

            string centre = Ro["BussinessCentre"].ToString();

            generic.Bookaslot(driver, extentTest, testName, testDataIteration, oppname, oppname);

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());



            //Navigating to activities and verifying language in payload

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["AsmLargeTourReques"].ToString());

            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadverificationwithASM(driver, extentTest, testName, testDataIteration, Ro["P1languagecodeUK"].ToString(), "Regus", emailid, Ro["TourASMtest3"].ToString());

            //CLOSE TOUR 
            generic.NavigateTourfrompayload(driver, extentTest, testName, testDataIteration);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Action1"].ToString());

            login.Logout(driver, extentTest, testName, testDataIteration);

            //}

            //catch (Exception e)

            //{



            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");

            //    Type thisType = this.GetType();

            //    object testCall = this;

            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            //}

        }




        public static IEnumerable<object[]> RTA_16797()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>

        ///[CRM-3363]-Verify whether the payload request for dailer contains the customer time zone information.

        /// </summary>

        [TestCategory("Regression"), TestCategory("GlobalP11"), TestCategory("1410"), TestCategory("0911"), TestCategory("GlobalP11A"), TestCategory("25Sep2020")]

        [TestProperty("TestcaseID", "RTA-16797")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_16797), DynamicDataSourceType.Method)]



        public void RTA_16797_Verifytimezoneindialerrequest(DataRow Ro)
        {
            //try { 

                       login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            string Date = DateTime.Today.ToString("dd-MM-yyyy");
            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            string emailid = "AutomatCompany" + Time + "@gmail.com";
            string now = System.DateTime.Now.ToString();

            //Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            //generic.DisbaleRTC(driver, extentTest, testName, testDataIteration);


            generic.EnterContactNumber(driver, extentTest, testName, testDataIteration, Ro["BussinessPhone"].ToString());
            generic.saveFooter(driver, extentTest, testName, testDataIteration);
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");

            //Creating new Opp
            //  generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "10";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, Ro["BussinessCentre"].ToString(), People);
            // string contactname = generic.CreateNewOpportunityBookATour(driver, extentTest, testName, testDataIteration,"London");

            string oppname = generic.GetOppID(driver, extentTest, testName, testDataIteration);


            //Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());

            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Selecting Activities and filling values for tour outcome
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField2"].ToString());
            generic.TourOutcome(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle1"].ToString(), Ro["P1noshow"].ToString(), Ro["P1noshowreason"].ToString());
            generic.SelectRelatedTab(driver, extentTest, testName, testDataIteration, Ro["P1RelatedField1"].ToString());

            //Navigating to activities and verifying language in payload
            generic.searchrequest(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle2"].ToString());
            generic.posttourcommrequestverify(driver, extentTest, testName, testDataIteration, Ro["ActivityTitle02"].ToString());
            generic.SelectingActiverequest(driver, extentTest, testName, testDataIteration);

            //Verifying Payload fields
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.payloadtimezoneverification(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString(), Ro["BussinessPhone"].ToString(), Ro["P1timezone"].ToString());

            login.Logout(driver, extentTest, testName, testDataIteration);
            //}
            //catch (Exception e)

            //{



            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");

            //    Type thisType = this.GetType();

            //    object testCall = this;

            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            //}
        }
        public static IEnumerable<object[]> RTA_25755()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        ///TEST [CRM-8319] -Verify comm request date format "Mm/dd/yyyy"  for the country
        /// </summary>

        [TestCategory("HotFix"), TestCategory("HFrerun")]

        [TestProperty("TestcaseID", "RTA_25755")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_25755), DynamicDataSourceType.Method)]

        public void RTA_25755_dateformat(DataRow Ro)
        {

            //try

            //{
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

            string Date = DateTime.Today.ToString("M/d/yyyy");
            string day = DateTime.Now.DayOfWeek.ToString();
            DateTime myDate = DateTime.Now;
            DateTime prevOne;
            if (day == "Friday")
            {
                prevOne = myDate.AddDays(3);
            }
            else if (day == "Saturday")
            {
                prevOne = myDate.AddDays(2);
            }
            
                prevOne = myDate.AddDays(1);
            

            DateTime Now = DateTime.Now;
            string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
            
           
            string furturedate = prevOne.ToString("M'/'dd'/'yyyy");
            string furturenewFormat = prevOne.ToString("yyyy-MM-dd");
            
            string centre = "Alberta, Calgary - Crowfoot Centre";
            string emailid = "AutomatCompany" + Time + "@gmail.com";

            string now = System.DateTime.Now.ToString();
          

            // Select Contacts entity and create new
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
           
            generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
            string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");

            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            //Creating new Opp
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
            string People = "13";
            generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, centre, People);

            string oppname = generic.GetOppID(driver, extentTest, testName, testDataIteration);

            //Verifying Opportunity status and Language
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());
            generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

            //Booking a tour

            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());
            string centre1 = centre;
            generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);
            generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

            //Navigating to activities and verifying language in payload
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["ActivityTitleDialer"].ToString());

            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            // generic.payloadverificationwithphone(driver, extentTest, testName, testDataIteration, Ro["Languagecode"].ToString(), Ro["regusbrand"].ToString(), emailid, Ro["payloadphone"].ToString(), Ro["TeamEmailqa"].ToString());
            generic.VerifyDailerScheduleddate(driver, extentTest, testName, testDataIteration, "", "", furturedate);
           
            generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["centervisitcommrequest"].ToString());
            //Verifying Payload fields.
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            // generic.payloadverificationwithphone(driver, extentTest, testName, testDataIteration, Ro["Languagecode"].ToString(), Ro["regusbrand"].ToString(), emailid, Ro["payloadphone"].ToString(), Ro["TeamEmailqa"].ToString());
            generic.VerifyDailerScheduleddate(driver, extentTest, testName, testDataIteration, "", "", furturedate);
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);

            //generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            //generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["centervisitcommrequest"].ToString());
            //generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            //// generic.payloadverificationwithphone(driver, extentTest, testName, testDataIteration, Ro["Languagecode"].ToString(), Ro["regusbrand"].ToString(), emailid, Ro["payloadphone"].ToString(), Ro["TeamEmailqa"].ToString());
            //generic.VerifyDailerScheduleddate(driver, extentTest, testName, testDataIteration, "", "", furturedate);

            //generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);

            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["P1searchtourrequest"].ToString());
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            // generic.payloadverificationwithphone(driver, extentTest, testName, testDataIteration, Ro["Languagecode"].ToString(), Ro["regusbrand"].ToString(), emailid, Ro["payloadphone"].ToString(), Ro["TeamEmailqa"].ToString());
            generic.VerifyDailerScheduleddate(driver, extentTest, testName, testDataIteration, "", "", furturedate);
            login.Logout(driver, extentTest, testName, testDataIteration);

            //}

            //catch (Exception e)

            //{



            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");

            //    Type thisType = this.GetType();

            //    object testCall = this;

            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            //}

        }

        public static IEnumerable<object[]> RTA_25784()
        {
            foreach (DataRow row in getTestCaseList("IWGSalesAgent"))
            {
                yield return new object[] { row };
            }
        }
        //TEST[CRM - 8319] -Verify comm request date format "dd/mm/yyyy"  for the country
        /// </summary>
        
        [TestCategory("HotFix"),TestCategory("HFrerun")]

        [TestProperty("TestcaseID", "RTA_25784")]


        [DataTestMethod]
        [DynamicData(nameof(RTA_25784), DynamicDataSourceType.Method)]

        public void RTA_25784_dateformat(DataRow Ro)
        {

            //try

            //{
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                string Date = DateTime.Today.ToString("dd/MM/yyyy");
                DateTime Now = DateTime.Now;
                string Time = Date + Convert.ToString(Now.Hour) + Convert.ToString(Now.Millisecond);
                DateTime myDate = DateTime.Now;
                DateTime prevOne = myDate.AddDays(1);
                string furturedate = prevOne.ToString("d/M/yyyy");
                string furturenewFormat = prevOne.ToString("yyyy-MM-dd");

                string centre = "London, HQ Moorgate";
                string emailid = "AutomatCompany" + Time + "@gmail.com";

                string now = System.DateTime.Now.ToString();

            Console.WriteLine(furturedate);
                // Select Contacts entity and create new
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", now, emailid, Ro["CompanyName"].ToString());
                string contactname = Element(driver, Control("CaseNumber", "GenericOld")).GetAttribute("innerText");



                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                //Creating new Opp
                generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
                generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());
                string People = "13";
                generic.CreateNewOpportunityExistingcontact(driver, extentTest, testDataIteration, testName, Ro["P1TestChannel"].ToString(), Ro["MajorSource"].ToString(), Ro["P1Minorsource"].ToString(), contactname, centre, People);

                string oppname = generic.GetOppID(driver, extentTest, testName, testDataIteration);




                //Verifying Opportunity status and Language

                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());

                generic.VerifyOppStatusandLanguage(driver, extentTest, testName, testDataIteration, Ro["UKEnglish"].ToString());

                //Booking a tour

                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Button"].ToString());

                string centre1 = centre;

                generic.Bookaslot(driver, extentTest, testName, testDataIteration, contactname, oppname);

                generic.SelectanyOpportunitybutton(driver, extentTest, testName, testDataIteration, Ro["Refresh"].ToString());



                //Navigating to activities and verifying language in payload

                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

                generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["ActivityTitleDialer"].ToString());

                //Verifying Payload fields.
                generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            // generic.payloadverificationwithphone(driver, extentTest, testName, testDataIteration, Ro["Languagecode"].ToString(), Ro["regusbrand"].ToString(), emailid, Ro["payloadphone"].ToString(), Ro["TeamEmailqa"].ToString());
            generic.VerifyDailerScheduleddate(driver, extentTest, testName, testDataIteration, "", "", furturedate);
                generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);

                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
                generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["centervisitcommrequest"].ToString());
                //Verifying Payload fields.
                generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            // generic.payloadverificationwithphone(driver, extentTest, testName, testDataIteration, Ro["Languagecode"].ToString(), Ro["regusbrand"].ToString(), emailid, Ro["payloadphone"].ToString(), Ro["TeamEmailqa"].ToString());
            generic.VerifyDailerScheduleddate(driver, extentTest, testName, testDataIteration, "", "", furturedate);
                generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);

            //generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            //generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["centervisitcommrequest"].ToString());
            //generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            //// generic.payloadverificationwithphone(driver, extentTest, testName, testDataIteration, Ro["Languagecode"].ToString(), Ro["regusbrand"].ToString(), emailid, Ro["payloadphone"].ToString(), Ro["TeamEmailqa"].ToString());
            //generic.VerifyDailerScheduleddate(driver, extentTest, testName, testDataIteration, "", "", furturedate);

            //generic.BacktooppotunityfromTourpage(driver, extentTest, testName, testDataIteration);
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.selectRequiredActivitty(driver, extentTest, testName, testDataIteration, Ro["P1searchtourrequest"].ToString());
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            generic.payloadnavigate(driver, extentTest, testName, testDataIteration);
            generic.HidingPureCloudNew(driver, extentTest, testName, testDataIteration);
            // generic.payloadverificationwithphone(driver, extentTest, testName, testDataIteration, Ro["Languagecode"].ToString(), Ro["regusbrand"].ToString(), emailid, Ro["payloadphone"].ToString(), Ro["TeamEmailqa"].ToString());
            generic.VerifyDailerScheduleddate(driver, extentTest, testName, testDataIteration, "", "", furturedate);
                login.Logout(driver, extentTest, testName, testDataIteration);

            //}

            //catch (Exception e)

            //{



            //    AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");

            //    Type thisType = this.GetType();

            //    object testCall = this;

            //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            //}

        }


        [TestCleanup]
        public void TestCleanup()
        {

            try
            {


                String teststatus = (TestContext.CurrentTestOutcome).ToString();
                String testExecutionKey = "CRM-9166";
                var testCaseId = GetType().GetMethod(TestContext.TestName).GetCustomAttributes(true).OfType<TestPropertyAttribute>().FirstOrDefault();
                String testkey = testCaseId.Value;
                String comment = generic.GenerateComment(teststatus);
                String json = generic.sendTestCaseJSON(testExecutionKey, testkey, teststatus, comment);
                generic.IntegrationTest(json, token);
                MyTestCleanup(TestContext.DataRow.Table.Rows.Count, "Yes");

            }



            catch (Exception e)
            {
                if (e.Message.Contains("Object reference not set to an instance of an object"))
                {
                    MyTestCleanup(0, "Yes");
                }
                else
                {
                    throw new Exception(e.Message);
                }
            }

            // killProcess("iexplore");
            // }


        }
    }
}
