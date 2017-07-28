using Microsoft.VisualStudio.TestTools.UnitTesting;
using AutoTestLib;
using System.Collections.Generic;
using System.Diagnostics;
using AutoTestLib.ExcelUtils;
using TechTalk.SpecFlow;

namespace UnitTest.AU
{
    [TestClass]
    public class UnitTest_ExcelUtils
    {
        [TestMethod]
        public void TestMethod_getColumnNumber()
        {
            string ColumnName;
            int ColumnNumber;

            ColumnName = "A";
            ColumnNumber = ExcelUtils.getColumnNumber(ColumnName);
            Assert.AreEqual(ColumnNumber, 1);


            ColumnName = "AA";
            ColumnNumber = ExcelUtils.getColumnNumber(ColumnName);
            Assert.AreEqual(ColumnNumber, 27);
        }

        [TestMethod]
        public void TestMethod_getLastRowNumber()
        {
            string excelFilePath = @"..\..\TestData\SampleData_1.xlsx";
            string sheetName = "Order";
            int columnNumber = 1;

            int totalRowNumber = ExcelUtils.getLastRowNumber(excelFilePath, sheetName, columnNumber);

            Assert.AreEqual(totalRowNumber, 39);


            excelFilePath = @"..\..\TestData\SampleData_2.xlsx";

           totalRowNumber = ExcelUtils.getLastRowNumber(excelFilePath, sheetName, columnNumber, 2);

            Assert.AreEqual(totalRowNumber, 40);

        }

        [TestMethod]
        public void TestMethod_ExcelToXML()
        {
            string XMLTemplateFilePath = @"..\..\TestData\ExcelToXML\books_template.xml";
            string excelFilePath = @"..\..\TestData\ExcelToXML\Books_ExcelToXML.xlsx";
            string sheetName = "TestData";
            string outputXMLFileFolder = @"..\..\TestData\ExcelToXML\output";
            int xPathRowNum = 1;
            int startCol = ExcelUtils.getColumnNumber("A");
            int endCol = ExcelUtils.getColumnNumber("G");
            int startRow = 2;
            int endRow = 3;
            int fileNameColNum = ExcelUtils.getColumnNumber("I");

            ExcelXMLUtils.ExcelToXML(XMLTemplateFilePath, excelFilePath, sheetName, outputXMLFileFolder, xPathRowNum, startCol, endCol, startRow, endRow, fileNameColNum);
        }

        [TestMethod]
        public void TestMethod_ExcelToXMLByConfig()
        {
            string excelFilePath = @"..\..\TestData\ExcelToXML\Books_ExcelToXML.xlsx";
            string configSheetName = "Configuration";

            string outputXMLFileFolder = ExcelXMLUtils.ExcelToXMLByConfig(excelFilePath, configSheetName);

            Assert.AreEqual(outputXMLFileFolder, @"..\..\TestData\ExcelToXML\output");
        }

        [TestMethod]
        public void TestMethod_XMLToExcelByConfig()
        {
            string excelFilePath = @"..\..\TestData\XMLToExcel\Books_XMLToExcel.xlsx";
            string configSheetName = "Configuration";

            ExcelXMLUtils.XMLToExcelByConfig(excelFilePath, configSheetName);            
        }

        [TestMethod]
        public void TestMethod_EreadKeyValuePair()
        {   
            var list = new List<KeyValuePair<string, string>>();
            string excelFilePath = @"..\..\TestData\ExcelComparison\Configuration.xlsx";
            string sheetName = "Configuration";

            list = ExcelUtils.readKeyValuePair(excelFilePath, sheetName, "B", "D");

          
            foreach (var item in list)
            {
                Debug.WriteLine("key is " + item.Key + ". value is " + item.Value);
            }

        }

        [TestMethod]
        public void TestMethod_CompareExcel()
        {
            string sourceFilePath = @"..\..\TestData\ExcelComparison\SourceBooks.xlsx";
            string sourceSheetName = "Books";
            int sourceStartRowNum = 2;
            int sourceEndRowNum = 4;
            string sourceOutputColName = "G";

            string targetFilePath = @"..\..\TestData\ExcelComparison\TargetBooks.xlsx";
            string targetSheetName = "Books";
            int targetStartRowNum = 2;
            int targetEndRowNum = 4;
            string targetOutputColName = "G";

            string configExcel = @"..\..\TestData\ExcelComparison\Configuration.xlsx";
            string configSheetName = "Configuration";
            string configSourceColName = "B";
            string configTargetColName = "D";


            ExcelComparison.compareExcelwithKeyColumn(sourceFilePath, sourceSheetName, sourceStartRowNum, sourceEndRowNum, sourceOutputColName,
                                                 targetFilePath, targetSheetName, targetStartRowNum, targetEndRowNum, targetOutputColName,
                                                 configExcel, configSheetName, configSourceColName, configTargetColName,
                                                 "LightGreen", "Tomato");
        }

        [TestMethod]
        public void MyTestMethod_saveSQLResultsInExcel()
        {            
            string connectionString = @"Server=BHPBIO-000228\SQLEXPRESS;Database=sampledb;Integrated Security=SSPI;";
            string queryString = @"select * from dbo.books";
            string excelFilePath = @"..\..\TestData\SQL2Excel\Books.xlsx";
            string sheetName = "books";

            SQL2Excel.saveSQLResultsInExcel(connectionString, queryString, excelFilePath, sheetName);
        }

        [TestMethod]
        public void TestMethod_readExcel2Table()
        {
            Excel2Table.readExcel2Table(3);
        }
    }
}
