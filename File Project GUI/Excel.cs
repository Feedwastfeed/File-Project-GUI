using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace File_Project_GUI
{
    class Excel
    {
        string path = "";
        Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public List<List<string>> readAll()
        {
            int countRows = ws.UsedRange.Rows.Count;
            int countColumns = ws.UsedRange.Columns.Count;
            Range range = (Range)ws.Range[ws.Cells[1, 1], ws.Cells[countRows, countColumns]];
            object[,] holder = range.Value2;
            List<List<string>> data = new List<List<string>>();
            
            for (int i = 1; i <= countRows; i++)
            {
                data.Add(new List<string>());
                for (int j = 1; j <= countColumns; j++)
                {
                    if(holder[i, j] == null)
                    {
                        data[i - 1].Add("");
                    }
                    else
                    {
                        data[i - 1].Add(holder[i, j].ToString());
                    }
                }
            }

            return data;
        }

        public void Close()
        {
            wb.Close();
        }


    }
}