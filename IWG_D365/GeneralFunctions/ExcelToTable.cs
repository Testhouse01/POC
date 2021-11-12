using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumCSharpMSTest.GeneralFunctions
{ 
   public class ExcelToTable
   {
    

    public static int row = 1;


    //public DataTable ExcelToDataTable(string fileName, string sheetName, string testCaseName)
    //{
    //        //open file and returns as Stream

           
    //    System.IO.FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
    //    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
    //    DataSet result = excelReader.AsDataSet();
    //    DataTableCollection table = result.Tables;
    //    stream.Close();

    //    string startupPath = Environment.CurrentDirectory;
    //    Console.WriteLine(startupPath);

    //    DataTable resultTable = table[sheetName];
    //    string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
    //    string actualPath = path.Substring(0, path.LastIndexOf("bin"));
    //        string projectPath = ExcelDataAccess.projectpath;
    //        string filePath = ExcelDataAccess.filePath;

    //        Console.WriteLine("ExcelTotable " + filePath);

    //        ExcelReaderFile exl = new ExcelReaderFile(fileName);
    //    int testCaseRowNumber1 = exl.getRowNumber(sheetName, 0, testCaseName + "_Start"); // check valuegobibo 17
    //    int testCaseRowNumber2 = exl.getRowNumber(sheetName, 0, testCaseName + "_End"); // gobibo 17

    //        Console.WriteLine("ExcelTotable testcase rownumber1 2 " + testCaseRowNumber1 + " " + testCaseRowNumber2);

    //        int testHeaderRowNumber = testCaseRowNumber1 + 1;
    //    int testDataStartRowNumber = testHeaderRowNumber + 1;
    //    int testDataEndRowNumber = testHeaderRowNumber;
    //    int counter = testDataStartRowNumber;
           
    //    while (!exl.getCellData(sheetName, 0, counter).Equals(""))
    //    {

    //        counter++;
    //        testDataEndRowNumber++;

    //    }
    //    testDataEndRowNumber = testDataEndRowNumber + 3;
    //        Console.WriteLine("ExcelTotable testDataEndRowNumber" + testDataEndRowNumber);
    //        try
    //    {

    //        for (int i = testCaseRowNumber2; i <= counter; i++)
    //        {

    //            resultTable.Rows[i].Delete();


    //        }

    //    }
    //    catch (Exception e)
    //    {

    //    }
    //    resultTable.AcceptChanges();
    //    try
    //    {

    //        for (int i = 0; i <= testCaseRowNumber1; i++)
    //        {

    //            resultTable.Rows[i].Delete();


    //        }

    //    }
    //    catch (Exception e)
    //    {

    //    }
    //    resultTable.AcceptChanges();

    //    try
    //    {
    //        foreach (DataColumn column in resultTable.Columns)
    //        {

    //            string cName = resultTable.Rows[0][column.ColumnName].ToString();
    //            if (!resultTable.Columns.Contains(cName) && cName != "")
    //            {
    //                column.ColumnName = cName;
    //            }

    //        }




    //        resultTable.Rows[0].Delete();
    //    }
    //    catch (Exception e)
    //    {

    //    }

    //    resultTable.AcceptChanges();
    //    stream.Close();
    //        int k = resultTable.Rows.Count;
    //        Console.WriteLine("resultTable count " + k);
    //        ////return
            
    //        return resultTable;
    //}

    
        
        
        
        
        //<Datacollection> dataCol = new List<Datacollection>();

    public DataTable PopulateInCollection(string sheetName,string filePath)
    {
        DataTable table = ExcelToDataTable(sheetName,filePath);

          
        for (int row = 1; row <= table.Rows.Count; row++)
        {
            for (int col = 0; col < table.Columns.Count; col++)
            {
                Datacollection dtTable = new Datacollection()
                {
                    rowNumber = row,
                    colName = table.Columns[col].ColumnName,
                    colValue = table.Rows[row - 1][col].ToString()
                };
                //dataCol.Add(dtTable);
            }
        }

            return table;

    }

        public DataTable ExcelToDataTable(string sheetName,string filePath)
        {
           
            //open file and returns as Stream
            System.IO.FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            //Createopenxmlreader via ExcelReaderFactory
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();
            //Get all the Tables
            DataTableCollection table = result.Tables;
            //Store it in DataTable
            DataTable resultTable = table[sheetName];


            foreach (DataColumn column in resultTable.Columns)
            {
                string cName = resultTable.Rows[0][column.ColumnName].ToString();
                if (!resultTable.Columns.Contains(cName) && cName != "")
                {
                    column.ColumnName = cName;
                }

            }

            resultTable.Rows[0].Delete();
            resultTable.AcceptChanges();
            stream.Close();
            ////return
            return resultTable;
        }


        public DataTable ExcelToDataTable1(string sheetName, string filePath)
        {

            //open file and returns as Stream
            System.IO.FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            //Createopenxmlreader via ExcelReaderFactory
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();
            stream.Close();
            //Get all the Tables
            DataTableCollection table = result.Tables;
            //Store it in DataTable
            DataTable resultTable = table[sheetName];


            foreach (DataColumn column in resultTable.Columns)
            {
                string cName = resultTable.Rows[0][column.ColumnName].ToString();
                if (!resultTable.Columns.Contains(cName) && cName != "")
                {
                    column.ColumnName = cName;
                }

            }

            resultTable.Rows[0].Delete();
            resultTable.AcceptChanges();
          
            ////return
            return resultTable;
        }

        public int rowCountFinder(string sheetName, string filePath )
    {
        DataTable table = ExcelToDataTable(sheetName,filePath);
        int rowCount = 0;
            rowCount = table.Rows.Count;
        return rowCount;
    }
    public void rowNumber(int rowN)

    {

        row = rowN;
    }
//    public string ReadData(string columnName)
//    {
//        try
//        {
//            //Retriving Data using LINQ to reduce much of iterations
//            string data = (from colData in dataCol
//                           where colData.colName == columnName && colData.rowNumber == row
//                           select colData.colValue).SingleOrDefault();

//            //var datas = dataCol.Where(x => x.colName == columnName && x.rowNumber == rowNumber).SingleOrDefault().colValue;
//            return data.ToString();
//        }
//        catch (Exception e)
//        {
//            return null;
//        }
//    }

}










public class Datacollection
{
    public int rowNumber { get; set; }
    public string colName { get; set; }
    public string colValue { get; set; }
}
}
