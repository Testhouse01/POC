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
    public class smoke : BaseClass
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


        

        [TestCleanup]
        public void TestCleanup()
        {

            try
            {


                String teststatus = (TestContext.CurrentTestOutcome).ToString();
                String testExecutionKey = "CRM-0000";
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
