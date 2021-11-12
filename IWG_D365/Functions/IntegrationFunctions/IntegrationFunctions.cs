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
    public class IntegrationFunctions : Helper
    {
        Reader reader = new Reader();
        public String currentStatus;
       
          
        
        /// <summary>
        /// Summary description for JiraUpdate 
        /// </summary

        public void IntegrationTest(String json, String token)
        {

            String accesstoken = token.Replace('"', ' ').Trim();
          //  Console.WriteLine(token);
            // API Initiation
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + accesstoken);


            // API Request
            var httpContent = new StringContent(json, Encoding.UTF8, "application/json");
            var res = client.PostAsync("https://xray.cloud.xpand-it.com/api/v1/import/execution", httpContent);

            // API Result
            try
            {
                res.Result.EnsureSuccessStatusCode();

                // Console.WriteLine("Response " + res.Result.Content.ReadAsStringAsync().Result + Environment.NewLine);

            }

            // Exception Handling
            catch (Exception ex)
            {
                // Console.WriteLine("Error " + res + " Error " + ex.ToString());
            }
            string result = "";
            //  Console.WriteLine("Response: {0}", result);
        }



        /// <summary>
        /// Generating Access token for Jira Update
        /// </summary

        public String AccessTokenGeneration(String json)
        {

            // API Initiation
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            // Console.WriteLine(json);


            // API Request
            var httpContent = new StringContent(json, Encoding.UTF8, "application/json");
            var res = client.PostAsync("https://xray.cloud.xpand-it.com/api/v1/authenticate", httpContent);


            String token = res.Result.Content.ReadAsStringAsync().Result;
            // API Result
            try
            {
                res.Result.EnsureSuccessStatusCode();

                //  Console.WriteLine("Response " + res.Result.Content.ReadAsStringAsync().Result + Environment.NewLine);

            }

            // Exception Handling
            catch (Exception ex)
            {
                // Console.WriteLine("Error " + res + " Error " + ex.ToString());
            }
            string result = "";
            //Console.WriteLine("Response: {0}", result);
            return token;
        }

     

        public String sendTestCaseJSON(String testExecutionKey, String key, String teststatus, String comment)
        {

            JObject sendObj = new JObject();
            sendObj.Add("testExecutionKey", testExecutionKey);
            JArray testCaseArray = new JArray();
            JObject testCaseObj = new JObject();
            testCaseObj.Add("testKey", key);
            testCaseObj.Add("status", teststatus);
            //  testCaseObj.Add("start", startTime);
            //testCaseObj.Add("finish", endTime);
            testCaseObj.Add("comment", comment);
            testCaseArray.Add(testCaseObj);
            sendObj.Add("tests", testCaseArray);
            // Console.WriteLine(key + "\n" + teststatus + "\n" + startTime + "\n" + endTime + "\n" + comment);
            String output = JsonConvert.SerializeObject(sendObj);
            return output;

        }

        public String accessTokenJSON(String clientid, String password)
        {
            JObject sendObj = new JObject();
            sendObj.Add("client_id", clientid);
            sendObj.Add("client_secret", password);
           // Console.WriteLine(clientid + "\n" + password + "\n");
            String output = JsonConvert.SerializeObject(sendObj);
            return output;
        }

        public String GenerateComment(String status)
        {
            String comment;
            if (status == "Failed")
            {
                comment = "The test case is Failed:- Please refer the test report for more details.";
            }
            else
            {
                comment = "Successfully completed test case execution.";
            }
            return comment;

        }

        public String gettingTestkey(String testname)

        {


            string resultString = Regex.Match(testname, @"\d+").Value;
            //Console.WriteLine(resultString);
            string str = "TEST-" + resultString;
            //Console.WriteLine(str);
            return str;


        }

     

        




    }
}











