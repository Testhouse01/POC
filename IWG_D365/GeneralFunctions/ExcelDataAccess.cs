using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace SeleniumCSharpMSTest.GeneralFunctions
{
    public class ExcelDataAccess
    {

        public static string sheetName;
        public static string filePath;
        public static string testName;
        public static string projectpath;


        public static void setExcelInfo(string sheetName1, string filePath1,  string projectpath1)
        {
            sheetName = sheetName1;
            filePath = filePath1;
           
            projectpath = projectpath1;

        }

        //public static List<DataRow> getData(System.Action methodWithParameters)
        //{
        //    DataTable runData = new DataTable();
        //    ExcelToTable d = new ExcelToTable();
        //    StackTrace stackTrace = new StackTrace();
        //    // string testName = stackTrace.GetFrame(1).GetMethod().Name;
        //    runData = d.PopulateInCollection(sheetName,filePath);
        //    int rowCount = d.rowCountFinder(sheetName,filePath);
        //    List<DataRow> rundatalist = new List<DataRow>();
        //    int itr = 0;
        //    for (int i = 1; i <= rowCount; i++)
        //    {


        //        d.rowNumber(i);
        //        if (d.ReadData("RunMode") == "Y")
        //        {
        //            rundatalist.Add(runData.Rows[i-1]);
        //            string iteration = itr.ToString();
        //            XmlDocument doc = new XmlDocument();
        //            doc.Load(projectpath + "testhousereport-config.xml");
        //            //XmlNode root = doc.DocumentElement;
        //            XmlNode root = doc.DocumentElement["configuration"];
        //            root.FirstChild.InnerText = testName;
        //            root.LastChild.InnerText = iteration;
        //            //XmlNode myNode = root.SelectSingleNode("descendant::iteration");
        //            //myNode.Value ="1";
        //            doc.Save(projectpath + "testhousereport-config.xml");
        //            methodWithParameters();
        //            itr++;
        //        }

        //    }


        //    return rundatalist;
        //}

        public static List<DataRow> getTestCaseData(string sheetName,string filePath)
        {
            DataTable runData = new DataTable();
            ExcelToTable d = new ExcelToTable();

            StackTrace stackTrace = new StackTrace();
            // string testName = stackTrace.GetFrame(1).GetMethod().Name;
            runData = d.PopulateInCollection(sheetName,filePath);
            int rowCount = d.rowCountFinder(sheetName,filePath);
            List<DataRow> rundatalist = new List<DataRow>();
            for (int i = 1; i <= rowCount; i++)
            {


              
                    rundatalist.Add(runData.Rows[i-1]);
                   
               

            }

            Console.WriteLine("this is getdata " + rundatalist.Count + " " + rowCount + " ");

            return rundatalist;

        }

        //public static DataRow CollectCurrentData(string id)
        //{
        //    DataTable runData = new DataTable();
        //    ExcelToTable d = new ExcelToTable();
        //    StackTrace stackTrace = new StackTrace();
        //    // string testName = stackTrace.GetFrame(1).GetMethod().Name;
        //    runData = d.PopulateInCollection( sheetName,filePath);
        //    int rowCount = d.rowCountFinder(sheetName,filePath);
        //    DataRow rundatalist = null;
        //    int itr = 0;

        //    for (int i = 1; i <= rowCount; i++)
        //    {


        //        d.rowNumber(i);
        //        if (d.ReadData("RunMode") == "Y" && id.Equals(d.ReadData("Id")))
        //        {
        //            rundatalist = runData.Rows[i - 1];
        //            break;
        //        }

        //    }


        //    return rundatalist;

        //}

        //public static DataRow CollectCurrentData()
        //{
        //    DataTable runData = new DataTable();
        //    ExcelToTable d = new ExcelToTable();
        //    StackTrace stackTrace = new StackTrace();
        //    // string testName = stackTrace.GetFrame(1).GetMethod().Name;
        //    runData = d.PopulateInCollection(sheetName,filePath);
        //    int rowCount = d.rowCountFinder(sheetName,filePath);
        //    DataRow rundatalist = null;
        //    int itr = 0;

        //    for (int i = 1; i <= rowCount; i++)
        //    {


        //        d.rowNumber(i);
        //        if (d.ReadData("RunMode") == "Y")
        //        {
        //            rundatalist = runData.Rows[i - 1];

              

        //            break;
        //        }

        //    }


        //    return rundatalist;

        //}


    }
}

