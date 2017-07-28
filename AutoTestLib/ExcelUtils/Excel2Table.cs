using System.Collections.Generic;
using TechTalk.SpecFlow;

namespace AutoTestLib.ExcelUtils
{
    public class Excel2Table
    {
        //Read data from a Excel file and put into a TechTalk.SpecFlow.Table class
        public static void readExcel2Table(int headerNum)
        {
            string[] header = new string[headerNum];

            for (int i = 0; i < headerNum; i++)
            {
                header[i] = "header" + i;
            }

            Table table = new Table(header);

            IDictionary<string, string> rows = new Dictionary<string, string>();

            for (int i = 0; i < headerNum; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    rows.Add(header[i], header[i] + "_row_" + j);
                   
                }
                
            }

            table.AddRow(rows);
            return;
        }

    }
}
