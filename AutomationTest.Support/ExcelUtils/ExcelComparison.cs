using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using OfficeOpenXml;
using System.IO;

namespace AutomationUtils
{
    public class ExcelComparison
    {
        //Compare two spreadsheets. Assumption is that source and target has one to one relationship on key column
        //Key column should be the first column in a config file
        //Firstly check key column, if key column matches then check other columns and lable color
        //If all columns match write "Match" in the output columns, otherwise write "Unmatched"
        //If key column is not matching, there is no color on that row, and write "Unmatched" in source sheet.        
        public static void compareExcelwithKeyColumn(string sourceFilePath, string sourceSheetName, int sourceStartRowNum, int sourceEndRowNum, string sourceOutputColName,
            string targetFilePath, string targetSheetName, int targetStartRowNum, int targetEndRowNum, string targetOutputColName, 
            string configExcel, string configSheetName, string configSourceColName, string configTargetColName,
            string strMatchColor = "LightGreen", string strUnmatchedColor = "Red")
        {
            Color matchColor = Color.FromName(strMatchColor);
            Color unmatchedColoer = Color.FromName(strUnmatchedColor);

            int sourceOutputColNum = ExcelUtils.getColumnNumber(sourceOutputColName);
            int targetOutputColNum = ExcelUtils.getColumnNumber(targetOutputColName);


            List <KeyValuePair<string, string>> columnNameList = ExcelUtils.readKeyValuePair(configExcel, configSheetName, configSourceColName, configTargetColName);

            List<KeyValuePair<int, int>> columnList = new List<KeyValuePair<int, int>>();
            foreach (var item in columnNameList)
            {
                columnList.Add(new KeyValuePair<int, int>(ExcelUtils.getColumnNumber(item.Key.ToString()), ExcelUtils.getColumnNumber(item.Value.ToString())));
            }
            

            var sourceExcel = new FileInfo(sourceFilePath);
            var targetExcel = new FileInfo(targetFilePath);

            using (ExcelPackage sourcePackage = new ExcelPackage(sourceExcel))
            {
                using (ExcelPackage targetPackage = new ExcelPackage(targetExcel))
                {

                    ExcelWorksheet sourceSheet = sourcePackage.Workbook.Worksheets[sourceSheetName];
                    ExcelWorksheet targetSheet = targetPackage.Workbook.Worksheets[targetSheetName];

                    sourceSheet.Calculate();
                    targetSheet.Calculate();

                    int sourceKeyColNum = columnList[0].Key;
                    int targetKeyColNum = columnList[0].Value;

                    string sourceCurrentKeyColValue = null;
                    int targetMatchedRowNum;

                    //loop every row from sourceSheet
                    for (int i = sourceStartRowNum; i <= sourceEndRowNum; i++)
                    {
                        bool matchWholeRow = false; 

                        //get current to-be-compared value from source sheet
                        if (null != sourceSheet.Cells[i, sourceKeyColNum].Value && (sourceSheet.Cells[i, sourceKeyColNum].Value.ToString().Length > 0))
                        {
                            //Comparing key column firstly, for one source record check all target records,                             
                            sourceCurrentKeyColValue = sourceSheet.Cells[i, sourceKeyColNum].Value.ToString();                             
                            targetMatchedRowNum = 0;
                            for (int j = targetStartRowNum; j <= targetEndRowNum; j++)
                            {
                                //if matching,  label source and target with color
                                if (null != targetSheet.Cells[j, targetKeyColNum].Value
                                        && (targetSheet.Cells[j, targetKeyColNum].Value.ToString().Length > 0)
                                        && targetSheet.Cells[j, targetKeyColNum].Value.ToString().Equals(sourceCurrentKeyColValue))
                                {
                                                                      
                                    sourceSheet.Cells[i, sourceKeyColNum].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    sourceSheet.Cells[i, sourceKeyColNum].Style.Fill.BackgroundColor.SetColor(matchColor);

                                    targetSheet.Cells[j, targetKeyColNum].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    targetSheet.Cells[j, targetKeyColNum].Style.Fill.BackgroundColor.SetColor(matchColor);


                                    targetMatchedRowNum = j;
                                    matchWholeRow = true;
                                    break;
                                }
                            }

                            //if key column matching,  Compare rest columns 
                            //source current row is i, target courrent row is targetMatchedRowNum
                            if (targetMatchedRowNum > 0)
                            {
                                //k is from 1 because the first column is key column
                                for (int k = 1; k < columnList.Count; k++)
                                {
                                    //if current sourceColNum > 0
                                    if (columnList[k].Key > 0)
                                    {
                                        string sourceStr = "";
                                        string targetStr = "";

                                        if (null != sourceSheet.Cells[i, columnList[k].Key].Value)
                                        {
                                            sourceStr = sourceSheet.Cells[i, columnList[k].Key].Value.ToString();                                           
                                        }

                                        if (null != targetSheet.Cells[targetMatchedRowNum, columnList[k].Value].Value)
                                        {
                                            targetStr = targetSheet.Cells[targetMatchedRowNum, columnList[k].Value].Value.ToString();
                                           
                                        }

                                        //Compare 
                                        if (sourceStr.Equals(targetStr))
                                        {                                           
       
                                            sourceSheet.Cells[i, columnList[k].Key].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            sourceSheet.Cells[i, columnList[k].Key].Style.Fill.BackgroundColor.SetColor(matchColor);

                                            targetSheet.Cells[targetMatchedRowNum, columnList[k].Value].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            targetSheet.Cells[targetMatchedRowNum, columnList[k].Value].Style.Fill.BackgroundColor.SetColor(matchColor);



                                        }
                                        else
                                        {
                                            sourceSheet.Cells[i, columnList[k].Key].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            sourceSheet.Cells[i, columnList[k].Key].Style.Fill.BackgroundColor.SetColor(unmatchedColoer);

                                            targetSheet.Cells[targetMatchedRowNum, columnList[k].Value].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            targetSheet.Cells[targetMatchedRowNum, columnList[k].Value].Style.Fill.BackgroundColor.SetColor(unmatchedColoer);

                                            matchWholeRow = false;


                                        }

                                    }

                                }

                            }

                            if (matchWholeRow)
                            {
                                sourceSheet.Cells[i, sourceOutputColNum].Value = "Match";
                                sourceSheet.Cells[i, sourceOutputColNum].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                sourceSheet.Cells[i, sourceOutputColNum].Style.Fill.BackgroundColor.SetColor(matchColor);

                                targetSheet.Cells[targetMatchedRowNum, targetOutputColNum].Value = "Match";
                                targetSheet.Cells[targetMatchedRowNum, targetOutputColNum].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                targetSheet.Cells[targetMatchedRowNum, targetOutputColNum].Style.Fill.BackgroundColor.SetColor(matchColor);
                            }
                            else
                            {
                                sourceSheet.Cells[i, sourceOutputColNum].Value = "Unmatched";
                                sourceSheet.Cells[i, sourceOutputColNum].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                sourceSheet.Cells[i, sourceOutputColNum].Style.Fill.BackgroundColor.SetColor(unmatchedColoer);

                                if (targetMatchedRowNum > 0)
                                {                                    
                                    targetSheet.Cells[targetMatchedRowNum, targetOutputColNum].Value = "Unmatched";
                                    targetSheet.Cells[targetMatchedRowNum, targetOutputColNum].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    targetSheet.Cells[targetMatchedRowNum, targetOutputColNum].Style.Fill.BackgroundColor.SetColor(unmatchedColoer);
                                }                                
                            }

                        }
                    }//loop sourceSheet 

                    sourcePackage.Save();
                    targetPackage.Save();

                }//using targetPackage     
            }//using sourcePackage     
        }
    }
}
