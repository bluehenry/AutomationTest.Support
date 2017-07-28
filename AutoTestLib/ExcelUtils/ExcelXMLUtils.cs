using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using OfficeOpenXml;

namespace AutoTestLib.ExcelUtils
{
    public class ExcelXMLUtils
    {
        public static void ExcelToXML(string XMLTemplateFilePath, string excelFilePath, string sheetName, string outputXMLFileFolder,
            int XPathRowNum, int startCol, int endCol, int startRow, int endRow, int fileNameColNum)
        {
            using (FileStream stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var package = new OfficeOpenXml.ExcelPackage(stream))
                {
                    ExcelWorkbook workBook = package.Workbook;
                    ExcelWorksheet workSheet = package.Workbook.Worksheets[sheetName];

                    string outputXmlFileFullPath = "";
                    string currentXPath = "";

                    for (int i = startRow; i <= endRow; i++)
                    {
                        outputXmlFileFullPath = Path.Combine(outputXMLFileFolder, workSheet.Cells[i, fileNameColNum].Value.ToString());

                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.Load(XMLTemplateFilePath);
                        XPathNavigator navigator = xmlDoc.CreateNavigator();



                        for (int j = startCol; j <= endCol; j++)
                        {
                            currentXPath = workSheet.Cells[XPathRowNum, j].Value.ToString();
                            XmlNode currentNode = xmlDoc.SelectSingleNode(currentXPath);


                            if (null == currentNode)
                            {
                                //If currentXPath is wrong, currentNode will be null                          
                                continue;
                            }

                            if (null != workSheet.Cells[i, j].Value)
                            {
                                //Even workSheet.Cells[i, j].Value.ToString().Length == 0, still process it
                                currentNode.InnerText = workSheet.Cells[i, j].Value.ToString();
                            }
                            else
                            {
                                //If there is no value, remove the node from xml doc.
                                currentNode.ParentNode.RemoveChild(currentNode);
                            }
                        }
                        xmlDoc.Save(outputXmlFileFullPath);
                    }
                }//using packge
            } //using stream
        }

        // read configuration from a Exel file and calls ExcelToXML() to generate XML files
        public static string ExcelToXMLByConfig(string excelFilePath, string configSheetName)
        {
            string XMLTemplateFilePath = "";
            string dataSheetName = "";
            int XPathRowNumber = 0;
            string XPathStartColName = "";
            string XPathEndColName = "";
            string fileNameColName = "";
            int dataStartRowNumber = 0;
            int dataEndRowNumber = 0;
            string outputXMLFileFolder = "";

            using (FileStream stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var package = new OfficeOpenXml.ExcelPackage(stream))
                {
                    ExcelWorkbook workBook = package.Workbook;
                    ExcelWorksheet configWorksheet = package.Workbook.Worksheets[configSheetName];

                    XMLTemplateFilePath = configWorksheet.Cells[2, 2].Value.ToString();
                    dataSheetName = configWorksheet.Cells[3, 2].Value.ToString();
                    XPathRowNumber = int.Parse(configWorksheet.Cells[4, 2].Value.ToString());
                    XPathStartColName = configWorksheet.Cells[5, 2].Value.ToString();
                    XPathEndColName = configWorksheet.Cells[6, 2].Value.ToString();
                    fileNameColName = configWorksheet.Cells[7, 2].Value.ToString();
                    dataStartRowNumber = int.Parse(configWorksheet.Cells[8, 2].Value.ToString());
                    dataEndRowNumber = int.Parse(configWorksheet.Cells[9, 2].Value.ToString());
                    outputXMLFileFolder = configWorksheet.Cells[10, 2].Value.ToString();

                    package.Dispose();

                }
            }

            ExcelToXML(XMLTemplateFilePath, excelFilePath, dataSheetName, outputXMLFileFolder,
                XPathRowNumber, ExcelUtils.getColumnNumber(XPathStartColName), ExcelUtils.getColumnNumber(XPathEndColName),
                dataStartRowNumber, dataEndRowNumber, ExcelUtils.getColumnNumber(fileNameColName));

            return outputXMLFileFolder;
        }

        //Read XML files from a folder, transfer them into a spreadsheet according to XPath
        //Deal with an XPath points to mutilple nodes: If all nodes have same value, output is the first node value. Otherwise output is all values concatnated together
        public static void XMLToExcel(string XMLFileFolder, string excelFilePath, string sheetName, int XPathRowNum, int startCol, int endCol,
            int fileNameColNum, int dataStartRowNum = 3)
        {
            //Get XML file path list
            List<string> XMLList = new List<string>();
            string[] fileList = Directory.GetFiles(XMLFileFolder);
            foreach (var file in fileList)
            {
                if (Path.GetExtension(file).Equals(".xml"))
                    XMLList.Add(file);
            }

            int currentRow = dataStartRowNum;

            var excelFile = new FileInfo(excelFilePath);
            using (var package = new OfficeOpenXml.ExcelPackage(excelFile))
            {
                ExcelWorkbook workBook = package.Workbook;
                ExcelWorksheet workSheet = package.Workbook.Worksheets[sheetName];

                foreach (var xmlfile in XMLList)
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    using (FileStream XMLStream = new FileStream(xmlfile, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        xmlDoc.Load(XMLStream);
                        XPathNavigator navigator = xmlDoc.CreateNavigator();
                        workSheet.Cells[currentRow, fileNameColNum].Value = xmlfile;

                        for (int currentCol = startCol; currentCol <= endCol; currentCol++)
                        {
                            if (null != workSheet.Cells[XPathRowNum, currentCol].Value && workSheet.Cells[XPathRowNum, currentCol].Value.ToString().Length > 0)
                            {
                                //XPath might point to multiple nodes
                                //If all nodes have same value, output is the first node value. Otherwise output is all values concatnated together
                                XmlNodeList selectedNodes = xmlDoc.SelectNodes(workSheet.Cells[XPathRowNum, currentCol].Value.ToString());

                                string firstNodesValue = "";
                                string currentNodesValue = "";
                                string allNodesValue = "";
                                bool allNodesValueAreEqual = true;

                                if (null != selectedNodes && selectedNodes.Count > 0)
                                {

                                    for (int nodeNum = 0; nodeNum < selectedNodes.Count; nodeNum++)
                                    {
                                        if (nodeNum == 0)
                                        {
                                            firstNodesValue = selectedNodes[nodeNum].InnerText;
                                            allNodesValue = selectedNodes[nodeNum].InnerText;
                                        }
                                        else
                                        {
                                            currentNodesValue = selectedNodes[nodeNum].InnerText;
                                            allNodesValue = allNodesValue + ";" + selectedNodes[nodeNum].InnerText;
                                            if (!currentNodesValue.Equals(firstNodesValue))
                                            {
                                                allNodesValueAreEqual = false;
                                            }
                                        }

                                        

                                    }//loop all nodes

                                }

                                if (allNodesValueAreEqual)
                                    workSheet.Cells[currentRow, currentCol].Value = firstNodesValue;
                                else
                                    workSheet.Cells[currentRow, currentCol].Value = allNodesValue;

                            }
                              
                        }
                        currentRow++;
                    }// using open a xml file
                }// end foreach, loop files
                package.Save();
            }            
        }

        // Read configruation from a Excel file and call XMLToExcel()
        public static void XMLToExcelByConfig(string excelFilePath, string configSheetName)
        {           
            string dataSheetName = "";
            int XPathRowNumber = 0;
            string XPathStartColName = "";
            string XPathEndColName = "";
            string fileNameColName = "";
            int dataStartRowNumber = 0;           
            string inXMLFileFolder = "";

            using (FileStream stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var package = new OfficeOpenXml.ExcelPackage(stream))
                {
                    ExcelWorkbook workBook = package.Workbook;
                    ExcelWorksheet configWorksheet = package.Workbook.Worksheets[configSheetName];
                    
                    dataSheetName = configWorksheet.Cells[2, 2].Value.ToString();
                    XPathRowNumber = int.Parse(configWorksheet.Cells[3, 2].Value.ToString());
                    XPathStartColName = configWorksheet.Cells[4, 2].Value.ToString();
                    XPathEndColName = configWorksheet.Cells[5, 2].Value.ToString();
                    fileNameColName = configWorksheet.Cells[6, 2].Value.ToString();
                    dataStartRowNumber = int.Parse(configWorksheet.Cells[7, 2].Value.ToString());                    
                    inXMLFileFolder = configWorksheet.Cells[8, 2].Value.ToString();

                    package.Dispose();

                }
            }

            XMLToExcel(inXMLFileFolder, excelFilePath, dataSheetName,
                XPathRowNumber, ExcelUtils.getColumnNumber(XPathStartColName), ExcelUtils.getColumnNumber(XPathEndColName),
               ExcelUtils.getColumnNumber(fileNameColName), dataStartRowNumber);

            return;
        }

    }
}
