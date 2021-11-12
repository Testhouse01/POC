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
    public class TestScript_POC : BaseClass
    {
        GenericFunctions generic = new GenericFunctions();
        LoginFunctions login = new LoginFunctions();
        IntegrationFunctions integration = new IntegrationFunctions();

        public static int maxTestRunsCount = 0;
        public static int maxTestRuns = maxTestRunsCount;
        public static String token;

        /// <summary>///
        /// Method to set report path.
        /// </summary>
        [AssemblyInitialize]
        public static void StartReport(TestContext test)
        {
            AssemblyInitialize();
        }

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
            var startTime = DateTime.Now.ToString("dd/MM/yyThh:mm:ss");
            TestInitialize();
            maxTestRuns = maxTestRunsCount;
        }

        public void InitializeForRerun()
        {
            TestInitialize();

        }
        #region DataReaderContext        

        public static IEnumerable<object[]> ReadDetails()
        {
            foreach (DataRow row in getTestCaseList("ITSalesUser"))
            {
                yield return new object[] { row };
            }
        }

        #endregion 

        #region ModelApps Positive Scenario


        /// <summary>
        /// To create Account in to Sample Model Application created in power apps
        /// Positive scenario
        /// </summary>
        //Pankaj
        [DataTestMethod]
        [DynamicData(nameof(ReadDetails), DynamicDataSourceType.Method)]
        public void CreateAccountandVerify(DataRow Ro)
        {
            try
            {   
                //Login into Application
                //login.Login(driver, extentTest, testName, testDataIteration, "https://orga4018de7.crm4.dynamics.com/main.aspx", Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, "https://orga4018de7.crm4.dynamics.com/main.aspx", "Pankaj.chaudhari@testhouse.net", "Vasanti*963");
                
                //Navigate to Sample Application MOdel App
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sample Application");
                
                //Navaigate to Accounts tab on left pane
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, "Accounts");
                
                ThinkTime(10);
                
                string Time = System.DateTime.Now.ToString();

                //Create a new Account
                string accountName = generic.CreateNewAccount(driver, extentTest, testDataIteration, testName, "Pran_Ltd", Time);

                ThinkTime(2);

                //Verify the account is created or not
                generic.VerifyAccounts(driver, extentTest, testName, testDataIteration, accountName);
                
                //Log Out of the application and close
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


        #endregion

        #region ModelApps-Negative Scenario

        /// <summary>
        /// To create Account in to Sample Application Model Application
        /// Negative scenario
        /// </summary>
        //Pankaj
        [DataTestMethod]
        //[TestProperty("TestcaseID", "RTA-5698")]
        [DynamicData(nameof(ReadDetails), DynamicDataSourceType.Method)]
        public void CreateAccountandVerifyNegative(DataRow Ro)
        {
            try
            {
                //Login into Application
                login.Login(driver, extentTest, testName, testDataIteration, "https://orga4018de7.crm4.dynamics.com/main.aspx", Ro["UserName"].ToString(), Ro["Password"].ToString());

                //Navigate to Sample Application MOdel App
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Sample Application");

                //Navaigate to Accounts tab on left pane
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, "Accounts");
                
                ThinkTime(10);
                
                string Time = System.DateTime.Now.ToString();
                
                //Create a new Account
                string accountName = generic.CreateNewAccount(driver, extentTest, testDataIteration, testName, "Pran International", Time);
                
                ThinkTime(2);

                //Verify the account is created or not
                accountName = "RandomAccount123";

                generic.VerifyAccounts(driver, extentTest, testName, testDataIteration, accountName);

                //Log Out of the application and close
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

        #endregion

        #region Customer Service


        /// <summary>
        /// To verify whether IT Sales user is not able to create Account
        /// </summary>
        //Pankaj
        [DataTestMethod]
        //[TestProperty("TestcaseID", "RTA-5698")]
        [DynamicData(nameof(ReadDetails), DynamicDataSourceType.Method)]
        public void CreateContactCaseandVerify(DataRow Ro)
        {
            try
            {
                //Login into Application
                //login.Login(driver, extentTest, testName, testDataIteration, "https://thdynamics365sandbox.crm4.dynamics.com/main.aspx", Ro["UserName"].ToString(), Ro["Password"].ToString());
                login.Login(driver, extentTest, testName, testDataIteration, "https://thdynamics365sandbox.crm4.dynamics.com/main.aspx", "Pankaj.chaudhari@testhouse.net", "Vasanti*963");
                ThinkTime(10);

                //Navigate to D-365 Customer Engagement Customer Service App
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Customer Service");

                //Navaigate to Contacts tab on left pane
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, "Contacts");

                //Create a new Contact
                string contact = generic.CreateNewContacts(driver, extentTest, testDataIteration, testName, Ro["CompanyName"].ToString(), "en-us", "crm.test3@regus.com", "crm.test3@regus.com", "12354689");

                ThinkTime(5);

                //Navaigate to Cases tab on left pane
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, "Cases");
                
                //Create a new Case
                string caseNumber = generic.CreateNewLSCCase(driver, extentTest, testName, testDataIteration, contact, "", "", "", "California, Century City - Avenue of the Stars", "");
                
                //Verify Case details
                generic.VerifyCase(driver, extentTest, testName, testDataIteration, caseNumber);

                //Log Out of the application and close
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

        #endregion

        #region Field Service


        /// <summary>
        /// To create Account in to D-365 Customer Engagement Field Service
        /// Create a work order
        /// Boook Resource
        /// Verif the resource
        /// </summary>
        [DataTestMethod]
        //[TestProperty("TestcaseID", "RTA-5698")]
        [DynamicData(nameof(ReadDetails), DynamicDataSourceType.Method)]
        public void CreateAccountandWorkOrder(DataRow Ro)
        {
            try
            {
                //Login into Application
                login.Login(driver, extentTest, testName, testDataIteration, "https://thdynamics365sandbox.crm4.dynamics.com/main.aspx", Ro["UserName"].ToString(), Ro["Password"].ToString());
                 

                //Navigate to Field Service Application
                generic.NavigateToMainDropDown(driver, extentTest, testName, testDataIteration, "Field Service");
                
                //Navigate to Work Orders from left pane
                generic.NavigateToEntity(driver, extentTest, testName, testDataIteration, "Work Orders");

                ThinkTime(10);

                string Time = System.DateTime.Now.ToString();

                //Create a account
                string accountName = generic.CreateNewAccountFieldService(driver, extentTest, testDataIteration, testName, "Pran_Ltd", Time);
                ThinkTime(2);

                //Create a work order
                string workorderID = generic.CreateNewWorkOrderFieldService(driver, extentTest, testDataIteration, testName, "Pran_Ltd", accountName);

                //verify the work order book candidate and verify
                generic.VerifyWorkOrder(driver, extentTest, testName, testDataIteration, workorderID);

                //Verify the account is created or not
                //generic.VerifyAccounts(driver, extentTest, testName, testDataIteration, accountName);
                //generic.SelectingOptionFromRelated(driver, extentTest, testName, testDataIteration, "RelatedTab", "Work Orders",accountName);
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

        #endregion

        #region TestCleanup

        [TestCleanup]
        public void TestCleanup()
        {

            try
            {


                String teststatus = (TestContext.CurrentTestOutcome).ToString();
                String testExecutionKey = "CRM-9166";
                var testCaseId = GetType().GetMethod(TestContext.TestName).GetCustomAttributes(true).OfType<TestPropertyAttribute>().FirstOrDefault();
                String testkey = testCaseId.Value;
                String comment = integration.GenerateComment(teststatus);
                String json = integration.sendTestCaseJSON(testExecutionKey, testkey, teststatus, comment);
                integration.IntegrationTest(json, token);
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


        }

        #endregion 

    }

}

