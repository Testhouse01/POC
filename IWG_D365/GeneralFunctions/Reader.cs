using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SeleniumCSharpMSTest.GeneralFunctions
{
    public class Reader
    {
        private List<Keywords> screens;
        private Keywords found = null;

        static string projectPath = Helper.SetProjectPath();
        string controlPath = projectPath + ConfigurationManager.AppSettings.Get("ControlsFilePath");


        public List<Keywords> ReadKeywordTable(DataTable table)
        {
            List<Keywords> screens = new List<Keywords>();

            string username = string.Empty;
            string password = string.Empty;

            foreach (DataRow row in table.Rows)
            {
                Keywords screen = new Keywords();

                string controlName = "";
                string propertyName = "";
                string propertyValue = "";

                if (row["ControlName"] != null)
                {
                    controlName = row["ControlName"].ToString();
                }
                if (row["PropertyName"] != null)
                {
                    propertyName = row["PropertyName"].ToString();
                }
                if (row["PropertyValue"] != null)
                {
                    propertyValue = row["PropertyValue"].ToString();
                }

                screen.ControlName = controlName;
                screen.PropertyName = propertyName;
                screen.PropertyValue = propertyValue;

                //adding keyword to list
                screens.Add(screen);

            }
            return screens;

        }

        /// <summary>
        /// Method to convert DataTable into a List
        /// </summary>
        /// <param name="table"></param>
        /// <param name="testcontext"></param>
        /// <returns></returns>
        public List<Keywords> ReadKeywordTable(DataTable table, TestContext testcontext)
        {
            List<Keywords> screens = new List<Keywords>();

            string username = string.Empty;
            string password = string.Empty;

            foreach (DataRow row in table.Rows)
            {
                Keywords screen = new Keywords();

                string controlName = "";
                string propertyName = "";
                string propertyValue = "";

                if (row["ControlName"] != null)
                {
                    controlName = row["ControlName"].ToString();
                }
                if (row["PropertyName"] != null)
                {
                    propertyName = row["PropertyName"].ToString();
                }
                if (row["PropertyValue"] != null)
                {
                    propertyValue = row["PropertyValue"].ToString();
                }

                screen.ControlName = controlName;
                screen.PropertyName = propertyName;
                screen.PropertyValue = propertyValue;

                //adding keyword to list
                screens.Add(screen);

            }
            return screens;

        }
        public static string dataFilePath(string file)
        {

            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actualPath = path.Substring(0, path.LastIndexOf("bin"));
            string projectPath = new Uri(actualPath).LocalPath;
            string filePath = projectPath + file;
            return filePath;
        }
        /// <summary>
        /// Method to Find Controls in the List
        /// </summary>
        /// <param name="controlName"></param>
        /// <param name="excelSheet"></param>
        /// <returns></returns>
        public Keywords FindControlinList(string controlName, string sheetName)
        { 
        ExcelToTable dTable = new ExcelToTable();
            DataTable table = dTable.ExcelToDataTable1(sheetName, dataFilePath("ObjectRepository\\"+sheetName +".xlsx"));

            for (int i = 0; i < ReadKeywordTable(table).Count; i++)
            {
                if (ReadKeywordTable(table)[i].ControlName == controlName)
                {
                    found = ReadKeywordTable(table)[i];
                    break;
                }
            }
            return found;
            
        }

        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        private TestContext testContextInstance;

    }
}
