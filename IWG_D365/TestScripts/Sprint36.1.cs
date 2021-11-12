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
    public class ITSale_Sprint36 : BaseClass
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


        public static IEnumerable<object[]> RTA_13907()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// To verify To verify the Recently created accounts And their Fields.
        /// </summary>
        [TestCategory("Opportunity"), TestCategory("IT Sales"), TestCategory("Priority1"), TestCategory("Regression"), TestCategory("Sprint3640")]
        [TestProperty("TestcaseID", "RTA-13907")]
        
[DataTestMethod]
[DynamicData(nameof(RTA_13907), DynamicDataSourceType.Method)]

        public void RTA13907_VerifyNewFieldsOpportunity(DataRow Ro)
        {
            //try
            //{

                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.CreateAndVerifyNewOpportunityITSales(driver, extentTest, testName, testDataIteration);
                generic.scrollUpPage(driver, extentTest, testName, testDataIteration);
                generic.VerifyNewOpportunityFieldsMouseHover(driver, extentTest, testName, testDataIteration);
                generic.FillMandatoryFieldsAndVerifyRecord(driver, extentTest, testName, testDataIteration);
                generic.saveFooter(driver, extentTest, testName, testDataIteration);

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
        public static IEnumerable<object[]> RTA_13925()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// Verify the IT Sales opportunity form fields of an opportunity record created via MS form         
        /// </summary>
        [TestCategory("Regression"), TestCategory("Priority1"), TestCategory("IT Sales"), TestCategory("Sprint3640"), TestCategory("25Nov")]
        [TestProperty("TestcaseID", "RTA-13925")]

        [DataTestMethod]
[DynamicData(nameof(RTA_13925), DynamicDataSourceType.Method)]
        public void RTA_13925_ITSalesOpportunityFormFieldsOfAnOpportunity(DataRow Ro)
        {
            //try
            //{
                generic.LaunchITSalesForm(driver, extentTest, testName, testDataIteration);
                string time = generic.GetSystemTimeInSec(driver, extentTest, testDataIteration, testName);
                string MSFormVerification = generic.SubmitingITSalesForm(driver, extentTest, testDataIteration, testName, Ro["CompanyName1"].ToString(), Ro["CustomerTitanid1"].ToString(), Ro["CustomerCont1"].ToString(), Ro["PersonEmail1"].ToString(), Ro["ContPersonPhno1"].ToString(), Ro["CentreCode1"].ToString(), Ro["AdditionalFeedback1"].ToString(), Ro["AddEmail1"].ToString(), time);
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, "Dashboards");
                generic.SelectFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Support User");

                //generic.verifyOppRecordCreatedInD365viaMSForm(driver, extentTest, testName, testDataIteration, Ro["CompanyName1"].ToString(), Ro["CustomerCont1"].ToString(), "Matched Contact");
                //SelectOpenITSalesOpportunities
                generic.verifyOppRecordCreatedInD365viaMSform(driver, extentTest, testName, testDataIteration, Ro["CompanyName1"].ToString());

                generic.DataPopulatedInOppFromITSalesForm(driver, extentTest, testName, testDataIteration, MSFormVerification, Ro["CustomerTitanid1"].ToString(), Ro["CustomerCont1"].ToString(), Ro["PersonEmail1"].ToString(), Ro["ContPersonPhno1"].ToString(), Ro["CentreCode1"].ToString(), Ro["AddEmail1"].ToString(), "Matched Contact");
                generic.Logout(driver, extentTest, testName, testDataIteration);
    //    }


    //        catch (Exception e)
    //        {

    //            AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
    //    Type thisType = this.GetType();
    //    object testCall = this;
    //    ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


    //}
}
        public static IEnumerable<object[]> RTA_14019()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// Verify the Welcome Email from Sitecore is sent to the customer Contact after 
        /// creating Opp record and system trigger automatic status update to "Silent" 
        /// for un-attended opportunity
        /// CRM-4575 &  CRM-4385
        /// </summary>
        [TestCategory("Regression"), TestCategory("Priority1"), TestCategory("Sprint36.1"), TestCategory("Opportunity"), TestCategory("Sprint3640")]
        [TestProperty("TestcaseID", "RTA-14019")]

        [DataTestMethod]
[DynamicData(nameof(RTA_14019), DynamicDataSourceType.Method)]
        public void RTA_14019_VerifySitecoreemailprocessandautomaticstatus(DataRow Ro)
        {
            //try
            //{
                // Login as IT Sales User and choose IT Sales
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                // Select Contacts entity and create new
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity1"].ToString());
                string time = generic.GetSystemTimeInSec(driver, extentTest, testDataIteration, testName);
                string contactname = generic.CreateNewContactITSaleswithlanguageandemail(driver, extentTest, testDataIteration, testName, "TestContact", time, "jeejams91@gmail.com", "");

                // Create new Opportunity with prev contact name
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.CreateNewOpportunityITSaleswithprevcontact(driver, extentTest, testDataIteration, testName, contactname, "New", "New Sale", "Long Term", "jeejams91@gmail.com");


                // Verify the process is at qualified
                generic.verifyprocessstageqaulify(driver, extentTest, testDataIteration, testName, contactname);


                // Verify status reason and status of the Opportunity
                generic.VerifyStatusReasonheader(driver, extentTest, testDataIteration, testName, "In Progress");
                generic.VerifyOpportunitystatus(driver, extentTest, testDataIteration, testName, "Open");

                // Go to Advance Find and verify the silent staus and minutes
                string parentwindow = generic.AdvancedFind(driver, extentTest, testName, testDataIteration);
                generic.FilterResultswithadditionalfilter(driver, extentTest, testName, testDataIteration, "Regus Settings", "key", "contains", "ITSales", parentwindow);

                // Logout from the application
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
        public static IEnumerable<object[]> RTA000()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// CRM-5498
        /// Verify the "Quantity" field on IT Sales Opportunity
        /// </summary>
        [TestCategory("Regression"), TestCategory("SprintUN")]
        [TestProperty("TestcaseID", "RTA-14882")]

        [DataTestMethod]
[DynamicData(nameof(RTA000), DynamicDataSourceType.Method)]
        public void RTA000_VerifyquantityfieldisdisplayedforOpportbuty(DataRow Ro)
        {
            try
            {



                // Login as IT Sales User and choose IT Sales
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());

                // Create new Opportunity 
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
                generic.ClickonNewButton(driver, extentTest, testName, testDataIteration, Ro["New Button"].ToString());


                // Verify the Quantity field



                // Logout from the application
                generic.Logout(driver, extentTest, testName, testDataIteration);
            }

            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            }
        }










        public static IEnumerable<object[]> RTA_14022()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// Contact data provided in the MS form doesn’t match a Dynamics contact record by Customer Titan ID – E-mail Contact person        
        /// </summary>
        [TestCategory("Regression"), TestCategory("25Nov"), TestCategory("IT Sales"), TestCategory("MSForm"), TestCategory("Sprint3640"), TestCategory("07Sep")]
        [TestProperty("TestcaseID", "RTA-14022")]

        [DataTestMethod]
[DynamicData(nameof(RTA_14022), DynamicDataSourceType.Method)]
        public void RTA_14022_ContactDataProvidedInMSFormDoesnMatchDynamicsContact(DataRow Ro)
        {
            try
            {
                generic.LaunchITSalesForm(driver, extentTest, testName, testDataIteration);
                string time = generic.GetSystemTimeInSec(driver, extentTest, testDataIteration, testName);
                string MSFormVerification = generic.SubmitingITSalesForm(driver, extentTest, testDataIteration, testName, Ro["CompanyName"].ToString(), Ro["CustomerTitanid"].ToString(), Ro["CustomerCont"].ToString(), Ro["PersonEmail"].ToString(), Ro["ContPersonPhno"].ToString(), Ro["CentreCode"].ToString(), Ro["AdditionalFeedback"].ToString(), Ro["AddEmail"].ToString(), time);
                //string MSFormVerification = Ro["CompanyName"].ToString();
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, "Dashboards");
                driver.Navigate().Refresh();
                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Support User");
                //generic.verifyOppRecordCreatedInD365viaMSForm(driver, extentTest, testName, testDataIteration, Ro["CompanyName"].ToString(), Ro["CustomerCont"].ToString(), "Unmatched Contact");
                generic.verifyOppRecordCreatedInD365viaMSform(driver, extentTest, testName, testDataIteration, Ro["CompanyName"].ToString());

                ThinkTime(10);
                generic.DataPopulatedInOppFromITSalesForm(driver, extentTest, testName, testDataIteration, MSFormVerification, Ro["CustomerTitanid"].ToString(), Ro["CustomerCont"].ToString(), Ro["PersonEmail"].ToString(), Ro["ContPersonPhno"].ToString(), Ro["CentreCode"].ToString(), Ro["AddEmail"].ToString(), "Unmatched Contact");
                ThinkTime(10);
                generic.verifyprocessstageqaulify1(driver, extentTest, testDataIteration, testName, "contactname");
                string parentwindow = generic.AdvancedFind(driver, extentTest, testName, testDataIteration);
                generic.FilterResultswithadditionalfilter(driver, extentTest, testName, testDataIteration, "Regus Settings", "key", "contains", "ITSales", parentwindow);
                generic.VerifyStatusReasonheader(driver, extentTest, testDataIteration, testName, "In Progress");
                generic.VerifyOpportunitystatus(driver, extentTest, testDataIteration, testName, "Open");
                generic.Logout(driver, extentTest, testName, testDataIteration);
            }


            catch (Exception e)
            {

                AddLog(driver, extentTest, testName, testDataIteration, "Fail", "Failed due to" + e.Message, "Test Failed");
                Type thisType = this.GetType();
                object testCall = this;
                ReRun(TestContext.TestName, ref maxTestRuns, maxTestRunsCount, extentTest, extentReport, thisType, testCall, "InitializeForRerun", "TestCleanup");


            }
        }

        public static IEnumerable<object[]> RTA_13999()
        {
            foreach (DataRow row in getTestCaseList("SharePointD3"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// TEST [CRM-4377] & [CRM-4381] Verify the functionality - Integration of “Move-Ins D-3” daily customers file as IT Sales opportunity and related phone call activity creation(automatically)
        /// </summary>
        [TestCategory("25Nov"), TestCategory("IT Sales"), TestCategory("Regression"), TestCategory("MoveinD3"), TestCategory("Sprint3640")]
        [TestProperty("TestcaseID", "RTA-13999")]

        [DataTestMethod]
[DynamicData(nameof(RTA_13999), DynamicDataSourceType.Method)]

        public void RTA_13999_ITSales_Opportunity_MoveInsD3(DataRow Ro)
        {

            // test record https://regusqa.crm4.dynamics.com/main.aspx?appid=79f0340e-d377-e911-a838-000d3ab18b89&pagetype=entityrecord&etn=opportunity&id=5ce0e13d-7d6f-ea11-a811-000d3ab8d1cd
            // sharePoint activities
            generic.loginSharePoint(driver, extentTest, testName, testDataIteration, uRL, Ro["SPUserName"].ToString(), Ro["SPPassword"].ToString());
            generic.selectSharePointFolder(driver, extentTest, testName, testDataIteration, Ro["FolderName1"].ToString(), Ro["FolderName2"].ToString(), Ro["FileName"].ToString());
            login.logoutSharePoint(driver, extentTest, testName, testDataIteration);

            // Navigating to D365 from SharePoint
            generic.sharePointToDynamics(driver, extentTest, testName, testDataIteration, uRL);

            //login to D365 application
            generic.Loginafterlogout2(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());

            // login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());

            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "New Opps - Move ins D-3");//"IT Sales Centre Opportunities");
            //readexcel.getExcelFile(driver, extentTest, testName, testDataIteration, Ro["FileName"].ToString());
            //generic.createdOppRecordsVerify((driver, extentTest, testName, testDataIteration, Ro["FileName"].ToString());
            generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
            string oppTopic = generic.SelectTitle(driver, extentTest, testName, testDataIteration);
            //string activityType = generic.verifyOppdataWithExcelData(driver, extentReport,  testName, testDataIteration);
            generic.VerifyBPFStageFieldITSalesBussinessProcess(driver, extentTest, testName, testDataIteration);
            string activityType = "E";
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.checkActivityCreation(driver, extentTest, testName, testDataIteration, activityType);

            generic.compareActivityDetailsWithOppDetails(driver, extentTest, testName, testDataIteration, oppTopic, activityType, "IT Sales Team", "AllData");

            if (activityType != "N")
            {
                generic.asssignOppRecordToLoggedInUser(driver, extentTest, testName, testDataIteration, "Me");
                generic.verifyAssignedUserInOpp(driver, extentTest, testName, testDataIteration, Ro["Owner"].ToString());
                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
                string activityNameCheck = generic.compareActivityDetailsWithOppDetails(driver, extentTest, testName, testDataIteration, oppTopic, activityType, Ro["Owner"].ToString(), "OwnerCheck");
                //string activityNameCheck = "100 Agency - India - Move-Ins M-3 Phone Call 3/30/2020 3:19 PM";
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, "Dashboards");
                generic.scrollMiddleDashboard(driver, extentTest, testName, testDataIteration, "Scroll");
                ThinkTime(3);
                generic.VerifyNewlyAssignedActivityAvailableInDashboard(driver, extentTest, testName, testDataIteration, activityNameCheck);
            }


            generic.HidingPureCloud(driver, extentTest, testName, testDataIteration);

            login.Logout(driver, extentTest, testName, testDataIteration);



        }




        public static IEnumerable<object[]> RTA_14000()
        {
            foreach (DataRow row in getTestCaseList("SharePointM3"))
            {
                yield return new object[] { row };
            }
        }

        /// <summary>
        /// TEST [CRM-4378] & [CRM-4381] Verify the functionality - Integration of “Move-Ins M-3” daily customers file as IT Sales opportunity and related phone call activity creation(automatically)
        /// </summary>
        [TestCategory("25Nov"), TestCategory("IT Sales"), TestCategory("Regression"), TestCategory("Sprint3640"), TestCategory("RF24-4-20"), TestCategory("RerunMay120"), TestCategory("Priority1")]
        [TestProperty("TestcaseID", "RTA-14000")]

        [DataTestMethod]
[DynamicData(nameof(RTA_14000), DynamicDataSourceType.Method)]

        public void RTA_14000_ITSales_Opportunity_MoveInsM3(DataRow Ro)
        {

            ////// test record https://regusqa.crm4.dynamics.com/main.aspx?appid=79f0340e-d377-e911-a838-000d3ab18b89&pagetype=entityrecord&etn=opportunity&id=5ce0e13d-7d6f-ea11-a811-000d3ab8d1cd
            ////// sharePoint activities
            ////generic.loginSharePoint(driver, extentTest, testName, testDataIteration, uRL, Ro["SPUserName"].ToString(), Ro["SPPassword"].ToString());
            ////generic.selectSharePointFolder(driver, extentTest, testName, testDataIteration, Ro["FolderName1"].ToString(), Ro["FolderName2"].ToString(), Ro["FileName"].ToString());
            ////login.logoutSharePoint(driver, extentTest, testName, testDataIteration);

            ////// Navigating to D365 from SharePoint
            ////generic.sharePointToDynamics(driver, extentTest, testName, testDataIteration, uRL);

            //////login to D365 application
            ////generic.Loginafterlogout(driver, extentTest, testName, testDataIteration, Ro["UserName"].ToString(), Ro["Password"].ToString());

            login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());

            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity"].ToString());
            generic.SelectValuesFromDropDown(driver, extentTest, testName, testDataIteration, "New Opps - Move ins M-3");//"IT Sales Centre Opportunities");
            readexcel.getExcelFile(driver, extentTest, testName, testDataIteration, Ro["FileName"].ToString());
            generic.SelectingActiveCell(driver, extentTest, testName, testDataIteration);
            string oppTopic = generic.SelectTitle(driver, extentTest, testName, testDataIteration);
            //string activityType = generic.verifyOppdataWithExcelData(driver, extentReport,  testName, testDataIteration);
            generic.VerifyBPFStageFieldITSalesBussinessProcess(driver, extentTest, testName, testDataIteration);
            string activityType = "P";
            generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());

            generic.checkActivityCreation(driver, extentTest, testName, testDataIteration, activityType);

            generic.compareActivityDetailsWithOppDetails(driver, extentTest, testName, testDataIteration, oppTopic, activityType, "IT Sales Team", "AllData");

            if (activityType != "N")
            {
                generic.asssignOppRecordToLoggedInUser(driver, extentTest, testName, testDataIteration, "Me");
                generic.verifyAssignedUserInOpp(driver, extentTest, testName, testDataIteration, Ro["Owner"].ToString());
                generic.SelectOpportunityTab(driver, extentTest, testName, testDataIteration, Ro["Tab"].ToString());
                string activityNameCheck = generic.compareActivityDetailsWithOppDetails(driver, extentTest, testName, testDataIteration, oppTopic, activityType, Ro["Owner"].ToString(), "OwnerCheck");
                //string activityNameCheck = "100 Agency - India - Move-Ins M-3 Phone Call 3/30/2020 3:19 PM";
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, "Dashboards");
                generic.scrollMiddleDashboard(driver, extentTest, testName, testDataIteration, "Scroll");
                ThinkTime(3);
                generic.VerifyNewlyAssignedActivityAvailableInDashboard(driver, extentTest, testName, testDataIteration, activityNameCheck);
            }


            // generic.HidingPureCloud(driver, extentTest, testName, testDataIteration);

            login.Logout(driver, extentTest, testName, testDataIteration);




        }



        ///// <summary>
        ///// TEST [CRM-4378] & [CRM-4381] Verify the functionality - Integration of “Move-Ins M-3” daily customers file as IT Sales opportunity and related phone call activity creation(automatically)
        ///// </summary>
        //[TestCategory("smoke"), TestCategory("IT Sales"), TestCategory("Regression")]
        //[TestProperty("TestcaseID", "RTA-14000")]
        //public static IEnumerable<object[]> RTA_5848_Data()
        //        {
        //            foreach (DataRow row in getTestCaseList("SharePoint"))
        //            {
        //                yield return new object[] { row };
        //            }
        //        }
        //[DataTestMethod]
        //[DataTestMethod]
//[DynamicData(nameof(RTA_5848_Data), DynamicDataSourceType.Method)]

        //public void excelDatareadingTest()
        //{
        //    readexcel.getExcelFile(driver, extentTest, testName, testDataIteration, Ro["FileName"].ToString());

        //}

        public static IEnumerable<object[]> RTA_14025()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// Verify the IT Sales App Dashboard views and each views filter condition contains "Opportunity 'Status Reason' not equal to "Silent""
        /// CRM-3401 & CRM-4385
        /// </summary>
        [TestCategory("Dashboard"), TestCategory("IT Sales"), TestCategory("27Aug"), TestCategory("Regression"), TestCategory("Sprint36.1"), TestCategory("Sprint3640")]
        [TestProperty("TestcaseID", "RTA-14025")]

        [DataTestMethod]
[DynamicData(nameof(RTA_14025), DynamicDataSourceType.Method)]

        public void RTA_14025_ITSales_Verifydashboardviewsandfilterconfitonstatus(DataRow Ro)
        {
            try
            {
                // Login as IT Sales User and select Dashboards
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());

            // Verify the Dashboard headers and dropdown values for IT inside Sales user
            generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Inside Sales User");
            generic.VerifyDashboardkpiheadersITSales(driver, extentTest, testName, testDataIteration, Ro["ITKPI1"].ToString(), Ro["ITKPI2"].ToString(), Ro["ITKPI3"].ToString(), Ro["ITKPI4"].ToString(), Ro["ITKPI5"].ToString(), "42", "3d", "9a", "");

            // Verify the option for Advanced Find status
            generic.scrollUpPagedashboard(driver, extentTest, testName, testDataIteration);
            generic.scrollUpPagedashboard(driver, extentTest, testName, testDataIteration);
            //generic.Verifyadvancefilterforallrecordsindashboardheader(driver, extentTest, testName, testDataIteration, Ro["ITinsideSalesusermore1"].ToString(), Ro["ITinsideSalesusermore2"].ToString(), Ro["ITinsideSalesusermore3"].ToString(), Ro["ITinsideSalesusermore4"].ToString(), Ro["ITKPI1"].ToString(), Ro["ITKPI2"].ToString(), Ro["ITKPI3"].ToString(), Ro["ITKPI4"].ToString(), "IT Inside Sales User");
            try
            {
                generic.Verifyadvancefilterforallrecordsindashboardheader1(driver, extentTest, testName, testDataIteration, Ro["ITinsideSalesusermore"].ToString(), Ro["ITinsideSalesusermore11"].ToString(), Ro["ITinsideSalesusermore12"].ToString(), Ro["ITinsideSalesusermore13"].ToString(), Ro["ITKPI1"].ToString(), Ro["ITKPI2"].ToString(), Ro["ITKPI3"].ToString(), Ro["ITKPI4"].ToString(), "IT Inside Sales User");

                // Select Unassigned Opps - Move-ins D-3 views for IT inside sales user and verify Advanced Find status
                generic.SelectsubmenuFromDropDown(driver, extentTest, testName, testDataIteration, "Unassigned Opps - Move-ins D-3");
                generic.Verifyadvancefindmultiplerecords1(driver, extentTest, testName, testDataIteration, Ro["ITinsideSalesUserMoree"].ToString(), "Unassigned Opps - Move-ins D-3", "IT Inside Sales User");

                // Select Unassigned Opps - Move-ins M-3 views for IT inside sales user and verify Advanced Find status
                generic.SelectsubmenuFromDropDown(driver, extentTest, testName, testDataIteration, "Unassigned Opps - Move-ins M-3");
                generic.Verifyadvancefindmultiplerecords1(driver, extentTest, testName, testDataIteration, Ro["ITinsideSalesUserMoreee"].ToString(), "Unassigned Opps - Move-ins M-3", "IT Inside Sales User");

                // Verify the headers and dropdown values for IT Sales support user
                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Support User");
                generic.VerifyDashboardkpiheadersITinsideSales(driver, extentTest, testName, testDataIteration, Ro["InsidesalesKPI1"].ToString(), Ro["InsidesalesKPI2"].ToString(), Ro["InsidesalesKPI3"].ToString(), Ro["InsidesalesKPI4"].ToString(), Ro["InsidesalesKPI5"].ToString());

                // Verify the option for Advanced Find status of all the views for IT Sales support user
                generic.scrollUpPagedashboard(driver, extentTest, testName, testDataIteration);
                generic.scrollUpPagedashboard(driver, extentTest, testName, testDataIteration);
                generic.Verifyadvancefilterforinsidesalesrecordsindashboardheader(driver, extentTest, testName, testDataIteration, Ro["ITSalesSupportMore1"].ToString(), Ro["ITSalesSupportMore2"].ToString(), Ro["ITSalesSupportMore3"].ToString(), Ro["ITSalesSupportMore4"].ToString(), Ro["InsidesalesKPI1"].ToString(), Ro["InsidesalesKPI2"].ToString(), Ro["InsidesalesKPI3"].ToString(), Ro["InsidesalesKPI4"].ToString(), "IT Sales Support User");
            }catch(Exception e)
            { }
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
        public static IEnumerable<object[]> RTA14792()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// Verify the IT Sales App Dashboard views and each views filter condition contains "Opportunity 'Status Reason' not equal to "Silent""
        /// </summary>
        [TestCategory("smoke"), TestCategory("Enterprise Sales"), TestCategory("Regression"), TestCategory("Sprint36.1"), TestCategory("Rerun-06062020"), TestCategory("Rerun-25062020"), TestCategory("SprintUN")]
        [TestProperty("TestcaseID", "RTA-14792")]

        [DataTestMethod]
[DynamicData(nameof(RTA14792), DynamicDataSourceType.Method)]

        public void RTA14792_ITSales_VerifydashboardheadersanddropdownoptionsforITsalesupportuser(DataRow Ro)
        {
            try
            {
                // Login as IT Sales User and select Dashboards
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());


                // Verify the headers and dropdown values for IT Sales support user
                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Support User");
                generic.VerifyDashboardkpiheadersITinsideSales(driver, extentTest, testName, testDataIteration, Ro["InsidesalesKPI1"].ToString(), Ro["InsidesalesKPI2"].ToString(), Ro["InsidesalesKPI3"].ToString(), Ro["InsidesalesKPI4"].ToString(), Ro["InsidesalesKPI5"].ToString());

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
        public static IEnumerable<object[]> RTA14790()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// Verify the IT Sales App Dashboard views and each views filter condition contains "Opportunity 'Status Reason' not equal to "Silent""
        /// </summary>
        [TestCategory("RerunMay9"), TestCategory("Enterprise Sales"), TestCategory("Regression"), TestCategory("Sprint36.1"), TestCategory("SprintUN")]
        [TestProperty("TestcaseID", "RTA-14790")]

        [DataTestMethod]
[DynamicData(nameof(RTA14790), DynamicDataSourceType.Method)]

        public void RTA14790_ITSales_VerifydashboardheadersanddropdownoptionsforITinsidesalesuser(DataRow Ro)
        {
            try
            {
                // Login as IT Sales User and select Dashboards
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());

                // Verify the Dashboard headers and dropdown values for IT inside Sales user
                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Inside Sales User");
                generic.VerifyDashboardkpiheadersITSales(driver, extentTest, testName, testDataIteration, Ro["ITKPI1"].ToString(), Ro["ITKPI2"].ToString(), Ro["ITKPI3"].ToString(), Ro["ITKPI4"].ToString(), Ro["ITKPI5"].ToString(), "42", "3d", "9a", "");

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

        public static IEnumerable<object[]> RTA15040()
        {
            foreach (DataRow row in getTestCaseList("ITSalesManager"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// CRM-5485
        /// Verify the IT INSIDE SALES USER dashboard and views
        /// </summary>
        [TestCategory("Sprint3640"), TestCategory("IT Sales"), TestCategory("Regression"), TestCategory("Priority1"), TestCategory("Sprint3640")]
        [TestProperty("TestcaseID", "RTA-15040")]

        [DataTestMethod]
[DynamicData(nameof(RTA15040), DynamicDataSourceType.Method)]

        public void RTA15040_VerifyITINSIDESALESMANAGERdashboard(DataRow Ro)
        {
            try
            {
                // Login as IT Sales User and select Dashboards
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());

                // Verify the Dashboard headers and dropdown values for IT inside Sales user
                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Manager");
                generic.ClickMoreManagerdashboardandseeallrecords(driver, extentTest, testName, testDataIteration, Ro["ITinsideManagerusermore1"].ToString(), Ro["ITKPI1"].ToString());

                // Verify the column headers for Open OPP
                generic.VerifyOpenOpps(driver, extentTest, testName, testDataIteration, Ro["Startdate"].ToString(), Ro["Enddate"].ToString(), Ro["Company"].ToString(), Ro["Pipeline"].ToString(), Ro["Brand"].ToString(), Ro["Columntest"].ToString(), Ro["Centre Number"].ToString(), Ro["Country"].ToString(), Ro["SaleType"].ToString(), Ro["NumberofUsers"].ToString(), Ro["TotalContract"].ToString(), Ro["Modifiedon"].ToString(), Ro["CreatedOn"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());

                // Verify the Dashboard headers and dropdown values for IT inside Sales user
                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Manager");

                // Select more for New IT Sales OPP
                generic.ClickMoreManagerdashboardandseeallrecords(driver, extentTest, testName, testDataIteration, Ro["ITinsideManagerusermore2"].ToString(), Ro["ITKPI2"].ToString());

                // Verify columns for New IT Sales OPP
                generic.VerifyNewITSalesOpps(driver, extentTest, testName, testDataIteration, Ro["Startdate"].ToString(), Ro["Enddate"].ToString(), Ro["Company"].ToString(), Ro["Pipeline"].ToString(), Ro["Brand"].ToString(), Ro["Columntest"].ToString(), Ro["Centre Number"].ToString(), Ro["Country"].ToString(), Ro["SaleType"].ToString(), Ro["NumberofUsers"].ToString(), Ro["TotalContract"].ToString(), Ro["Modifiedon"].ToString(), Ro["CentreOrgUn"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());

                //Verify the Dashboard headers and dropdown values for IT inside Sales user

                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Manager");

                // Select more for Activities
                generic.ClickMoreManagerdashboardandseeallrecords(driver, extentTest, testName, testDataIteration, Ro["ITinsideManagerusermore3"].ToString(), Ro["ITKPI3"].ToString());

                generic.ActivityError1(driver, extentTest, testName, testDataIteration, Ro["ITinsideManagerusermore3"].ToString(), Ro["ITKPI3"].ToString());


                // Verify columns for Activities
                generic.VerifyActivitiesinall(driver, extentTest, testName, testDataIteration, Ro["Duedate"].ToString(), Ro["Activitytype"].ToString(), Ro["regardingobjectid"].ToString(), Ro["Brandid"].ToString(), Ro["Owner"].ToString(), Ro["Busphone"].ToString(), Ro["email1"].ToString(), Ro["fullname"].ToString(), Ro["Phone"].ToString(), Ro["Contactemail"].ToString(), Ro["Contactfullname"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());
                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Manager");


                generic.scrollMiddleDown(driver, extentTest, testName, testDataIteration);

                // Select more for Open Opps chart
                generic.ClickMoreandseeallrecordsforchart(driver, extentTest, testName, testDataIteration, Ro["ITinsideManagerusermore4"].ToString(), Ro["ITKPI4"].ToString());

                // Verify columns for Open Opps chart
                generic.VerifyOpenOpps(driver, extentTest, testName, testDataIteration, Ro["Startdate"].ToString(), Ro["Enddate"].ToString(), Ro["Company"].ToString(), Ro["Pipeline"].ToString(), Ro["Brand"].ToString(), Ro["Columntest"].ToString(), Ro["Centre Number"].ToString(), Ro["Country"].ToString(), Ro["SaleType"].ToString(), Ro["NumberofUsers"].ToString(), Ro["TotalContract"].ToString(), Ro["Modifiedon"].ToString(), Ro["CreatedOn"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());
                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Manager");

                generic.scrollMiddleDown(driver, extentTest, testName, testDataIteration);
                // Select more for Closed Opps chart
                generic.ClickMoreandseeallrecordsforchart(driver, extentTest, testName, testDataIteration, Ro["ITinsideManagerusermore5"].ToString(), Ro["ITKPI5"].ToString());

                // Verify columns for Closed Opps chart
                generic.VerifyClosedOpps(driver, extentTest, testName, testDataIteration, Ro["ActualClose"].ToString(), Ro["CreatedOn"].ToString(), Ro["Topic"].ToString(), Ro["Potential"].ToString(), Ro["Country"].ToString(), Ro["SaleType"].ToString(), Ro["Columntest"].ToString(), Ro["Field9"].ToString(), Ro["ActualRevenue"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());
                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Manager");
                generic.scrollMiddleDown(driver, extentTest, testName, testDataIteration);

                // Select more for Won Opps chart
                generic.ClickMoreandseeallrecordsforchart(driver, extentTest, testName, testDataIteration, Ro["ITinsideManagerusermore6"].ToString(), Ro["ITKPI6"].ToString());

                // Verify columns for Won Opps chart
                generic.VerifyClosedOpps(driver, extentTest, testName, testDataIteration, Ro["ActualClose"].ToString(), Ro["CreatedOn"].ToString(), Ro["Topic"].ToString(), Ro["Potential"].ToString(), Ro["Country"].ToString(), Ro["SaleType"].ToString(), Ro["Columntest"].ToString(), Ro["Field9"].ToString(), Ro["ActualRevenue"].ToString());

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
        public static IEnumerable<object[]> RTA15039()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }
        /// <summary>
        /// CRM-5485
        /// Verify the IT INSIDE SALES USER dashboard and views
        /// </summary>
        [TestCategory("Sprint3640"), TestCategory("IT Sales"), TestCategory("Regression"), TestCategory("Priority1"), TestCategory("Sprint3640")]
        [TestProperty("TestcaseID", "RTA-15039")]

        [DataTestMethod]
[DynamicData(nameof(RTA15039), DynamicDataSourceType.Method)]

        public void RTA15039_VerifyITSalesSupportUserdashboard(DataRow Ro)
        {
            try
            {
                // Login as IT Sales User and select Dashboards
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());

                // Verify the Dashboard headers and dropdown values for IT inside Sales user
                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Support User");
                generic.ClickMoreandseeallrecords2(driver, extentTest, testName, testDataIteration, Ro["ITSalesSupportMore1"].ToString(), Ro["InsidesalesKPI1"].ToString());

                // Verify the column headers for UnAssigned Opps
                generic.VerifyUnassignedOppscolumns(driver, extentTest, testName, testDataIteration, Ro["CreatedOn"].ToString(), Ro["Country"].ToString(), Ro["Company"].ToString(), Ro["Brand"].ToString(), Ro["Startdate"].ToString(), Ro["Enddate"].ToString(), Ro["SaleType"].ToString(), Ro["Recommended"].ToString(), Ro["Languageheader"].ToString(), Ro["NumberofUsers"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());

                // Verify the Dashboard headers and dropdown values for IT inside Sales user
                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Support User");

                // Select more for Open Opps
                generic.ClickMoreandseeallrecords2(driver, extentTest, testName, testDataIteration, Ro["ITSalesSupportMore2"].ToString(), Ro["InsidesalesKPI2"].ToString());

                // Verify columns for Open Opps
                generic.VerifyOpenOpps(driver, extentTest, testName, testDataIteration, Ro["Startdate"].ToString(), Ro["Enddate"].ToString(), Ro["Company"].ToString(), Ro["Pipeline"].ToString(), Ro["Brand"].ToString(), Ro["Column5"].ToString(), Ro["Centre Number"].ToString(), Ro["Country"].ToString(), Ro["SaleType"].ToString(), Ro["NumberofUsers"].ToString(), Ro["TotalContract"].ToString(), Ro["Modifiedon"].ToString(), Ro["CreatedOn"].ToString());

                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());

                //Verify the Dashboard headers and dropdown values for IT inside Sales user

                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Support User");

                //// Select more for Activities
                generic.scrollMiddleDown(driver, extentTest, testName, testDataIteration);

                generic.ClickMoreandseeallrecords2(driver, extentTest, testName, testDataIteration, Ro["ITSalesSupportMore3"].ToString(), Ro["InsidesalesKPI3"].ToString());

                // Verify columns for Activities
                generic.ActivityError2(driver, extentTest, testName, testDataIteration, Ro["ITSalesSupportMore3"].ToString(), Ro["InsidesalesKPI3"].ToString());

                generic.VerifyActivitiesinall(driver, extentTest, testName, testDataIteration, Ro["Duedate"].ToString(), Ro["Activitytype"].ToString(), Ro["regardingobjectid"].ToString(), Ro["Brandid"].ToString(), Ro["Owner"].ToString(), Ro["Busphone"].ToString(), Ro["email1"].ToString(), Ro["fullname"].ToString(), Ro["Phone"].ToString(), Ro["Contactemail"].ToString(), Ro["Contactfullname"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());
                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Support User");
                generic.scrollMiddleDown(driver, extentTest, testName, testDataIteration);

                // Select more for Open Opps chart
                generic.ClickMoreandseeallrecordsforchart(driver, extentTest, testName, testDataIteration, Ro["ITSalesSupportMore4"].ToString(), Ro["InsidesalesKPI4"].ToString());

                // Verify columns for Open Opps chart
                generic.VerifyOpenOpps(driver, extentTest, testName, testDataIteration, Ro["Startdate"].ToString(), Ro["Enddate"].ToString(), Ro["Company"].ToString(), Ro["Pipeline"].ToString(), Ro["Brand"].ToString(), Ro["Column5"].ToString(), Ro["Centre Number"].ToString(), Ro["Country"].ToString(), Ro["SaleType"].ToString(), Ro["NumberofUsers"].ToString(), Ro["TotalContract"].ToString(), Ro["Modifiedon"].ToString(), Ro["CreatedOn"].ToString());
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());
                generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Support User");
                generic.scrollDownDashboard(driver, extentTest, testName, testDataIteration);

                // Select more for Closed Opps chart
                generic.ClickMoreandseeallrecordsforchart(driver, extentTest, testName, testDataIteration, Ro["ITSalesSupportMore5"].ToString(), Ro["InsidesalesKPI5"].ToString());

                // Verify columns for Closed Opps chart
                generic.VerifyClosedOpps(driver, extentTest, testName, testDataIteration, Ro["ActualClose"].ToString(), Ro["CreatedOn"].ToString(), Ro["Topicfield1"].ToString(), Ro["Potential"].ToString(), Ro["Country"].ToString(), Ro["SaleType"].ToString(), Ro["Column5"].ToString(), Ro["Field9"].ToString(), Ro["ActualRevenue"].ToString());

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




























        public static IEnumerable<object[]> RTA_14030()
        {
            foreach (DataRow row in getTestCaseList("ITSalesManager"))
            {
                yield return new object[] { row };
            }
        }

        /// <summary>
        /// Verify the IT Sales Manager Dashboard views and ech views filter condition contains "Opportunity 'Status Reason' not equal to "Silent""
        /// CRM-3401 & CRM-4385
        /// </summary>
        [TestCategory("27Aug"), TestCategory("IT Sales"), TestCategory("Regression"), TestCategory("Sprint36.1"), TestCategory("Dashboard"), TestCategory("Sprint3640")]
        [TestProperty("TestcaseID", "RTA-14030")]

        [DataTestMethod]
[DynamicData(nameof(RTA_14030), DynamicDataSourceType.Method)]

        public void RTA_14030_ITSalesMngrVrfydashboardviewsNdfilterconfitonstatus(DataRow Ro)
        {

            try
            {
                //Login as IT Sales Manager and select Dashboards
                login.Login(driver, extentTest, testName, testDataIteration, uRL, Ro["UserName"].ToString(), Ro["Password"].ToString());
            generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, Ro["Sales"].ToString());
            generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, Ro["Entity0"].ToString());

            // Verify the Dashboard headers and dropdown values for IT  Sales Manager user
            generic.SelectValuesFromDashboardDropDown(driver, extentTest, testName, testDataIteration, "IT Sales Manager");
            generic.VerifyDashboardkpiheadersITSalesManager(driver, extentTest, testName, testDataIteration, Ro["KPI1"].ToString(), Ro["KPI2"].ToString(), Ro["KPI3"].ToString(), Ro["KPI1"].ToString(), Ro["KPI5"].ToString(), Ro["KPI6"].ToString(), "a6", "34", "f5");

            // Verify all drop down options for Open It sales and verify Advanced Find status
            //generic.Manager_Verifyadvancefilterforopenrecordsindashboardheader(driver, extentTest, testName, testDataIteration, Ro["ITinsideSalesusermore1"].ToString(), Ro["ITinsideSalesusermore2"].ToString(), Ro["ITinsideSalesusermore3"].ToString(), Ro["ITinsideSalesusermore4"].ToString(), Ro["ITinsideSalesusermore5"].ToString(), Ro["ITinsideSalesusermore6"].ToString(), Ro["KPI1"].ToString(), Ro["KPI2"].ToString(), Ro["KPI3"].ToString(), Ro["KPI4"].ToString(), Ro["KPI5"].ToString(), Ro["KPI6"].ToString(), "IT Sales Manager");
            generic.Manager_Verifyadvancefilterforopenrecordsindashboardheader(driver, extentTest, testName, testDataIteration, Ro["ITSalesManagerMore1"].ToString(), Ro["ITSalesManagerMore2"].ToString(), Ro["ITSalesManagerMore3"].ToString(), Ro["ITSalesManagerMore4"].ToString(), Ro["ITSalesManagerMore5"].ToString(), Ro["ITinsideSalesusermore6"].ToString(), Ro["KPI1"].ToString(), Ro["KPI2"].ToString(), Ro["KPI3"].ToString(), Ro["KPI4"].ToString(), Ro["KPI5"].ToString(), Ro["KPI6"].ToString(), "IT Sales Manager");

            // Verify all drop down options for New It sales and verify Advanced Find status
            generic.Manager_Verifyadvancefilterfornewrecordsindashboardheader(driver, extentTest, testName, testDataIteration, Ro["ITSalesManagerMore11"].ToString(), Ro["ITSalesManagerMore12"].ToString(), Ro["ITSalesManagerMore13"].ToString(), Ro["ITSalesManagerMore14"].ToString(), Ro["ITSalesManagerMore15"].ToString(), Ro["ITinsideSalesusermore6"].ToString(), Ro["KPI1"].ToString(), Ro["KPI2"].ToString(), Ro["KPI3"].ToString(), Ro["KPI4"].ToString(), Ro["KPI5"].ToString(), Ro["KPI6"].ToString(), "IT Sales Manager");

            // Verify Advance Filter staus for Activities and chart
            generic.Verifyadvancefilterforactivityandchartrecords(driver, extentTest, testName, testDataIteration, Ro["ITSalesManagerMore21"].ToString(), Ro["ITSalesManagerMore31"].ToString(), Ro["KPI3"].ToString(), Ro["KPI4"].ToString(), "IT Sales Manager");

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
        }


    }
}
