using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
namespace ExcelClient
{
    class Writer
    {
        private System.Data.DataTable DataTable;
        private Workbook workbook;
        private Worksheet worksheet;
        private Application excel;

        public Writer(System.Data.DataTable dataTable,string sheetName, string tableName, string path)
        {
            //construct
            DataTable = dataTable;
            excel = new Application
            {
                Visible = false,
                DisplayAlerts = false
            };
            workbook = excel.Workbooks.Add(Type.Missing);
            worksheet = (Worksheet) workbook.ActiveSheet;
            worksheet.Name = sheetName;
            // merge cells
            //worksheet.Range[worksheet.Cells[1,1], worksheet.Cells[1, 8]].Merge();
            // Table Name
            worksheet.Cells[1, 1] = tableName;
            worksheet.Cells.Font.Size = 15;
            
            
            AddTable(2,1);
            workbook.SaveAs(path); ;
            
        }

        public void CloseExcel()
        {
            workbook.Close();
            excel.Quit();
            // get rid of these so you can construct again 
            DataTable = null;
            workbook = null;
            worksheet = null;
            excel = null;
    }

        public void AddTable(int startX, int startY)
        {
            //borders
            Range cellRange = worksheet.Range[worksheet.Cells[startX, startY], worksheet.Cells[(startX - 1) + DataTable.Rows.Count, DataTable.Columns.Count]];
            Borders border = cellRange.Borders;
            border.LineStyle = XlLineStyle.xlContinuous;
            border.Weight = 2d;
            
            // header
            Range headerRange = worksheet.Range[worksheet.Cells[startX, startY], worksheet.Cells[startX , DataTable.Columns.Count]];
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.Aqua);
            
            // data
            int xCounter = startX;
            int yCounter = startY;
            foreach (DataRow row in DataTable.Rows)
            {
                foreach (DataColumn column in DataTable.Columns)
                {
                    var cell = row[column];
                    worksheet.Cells[xCounter,yCounter] = row[column];
                    yCounter++;
                }

                yCounter = startY;
                xCounter++;

            }
            cellRange.EntireColumn.AutoFit();

            //celLRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, DataTable.Columns.Count]];
        }

    }
}
