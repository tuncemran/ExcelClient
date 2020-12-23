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
            worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 14]].Merge();
            worksheet.Cells[1, 1] = tableName;
            worksheet.Cells.Font.Size = 15;

            worksheet.Cells[2, 2] = "ORGANİZASYON";
            worksheet.Cells[2, 3] = "SÜREÇ";
            // cell merger
            // kişsel veri
            worksheet.Range[worksheet.Cells[2, 4], worksheet.Cells[2, 9]].Merge();
            worksheet.Cells[2, 4] = "KİŞİSEL VERİ";

            worksheet.Cells[2, 10] = "SAKLAMA ve İMHA";
            // aktarma
            worksheet.Range[worksheet.Cells[2, 11], worksheet.Cells[2, 12]].Merge();
            worksheet.Cells[2, 11] = "AKTARMA";
            //alınan güvenlik tedbirleri
            
            worksheet.Cells[2, 13] = "ALINAN GÜVENLİK TEDBİRLERİ";
            worksheet.Range[worksheet.Cells[2, 13], worksheet.Cells[2, 14]].Columns.AutoFit();
            worksheet.Range[worksheet.Cells[2, 13], worksheet.Cells[2, 14]].Merge();

            // add columns
            
            int xIndex = 1;
            foreach (DataColumn column in DataTable.Columns)
            {
                worksheet.Cells[3, xIndex] = column.ColumnName;
                xIndex++;
            }
            worksheet.Range[worksheet.Cells[3, 1], worksheet.Cells[2, 14]].Columns.AutoFit();
            AddTable(4,1);
            AddStyleToTable();
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

            // data
            int xCounter = startX;
            int yCounter = startY;
            foreach (DataRow row in DataTable.Rows)
            {
                foreach (DataColumn column in DataTable.Columns)
                {
                    worksheet.Cells[xCounter,yCounter] = row[column];
                    yCounter++;
                }

                
                worksheet.Range[worksheet.Cells[xCounter, 1], worksheet.Cells[xCounter + 1, DataTable.Columns.Count]].Rows.Style.VerticalAlignment =
                    XlVAlign.xlVAlignTop;

                yCounter = startY;
                xCounter++;

            }

        }

        public void AddStyleToTable()
        {
            Range cellRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[3 + DataTable.Rows.Count, DataTable.Columns.Count]];
            cellRange.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            //cellRange.Columns.AutoFit();
            cellRange.Rows.AutoFit();
            Borders border = cellRange.Borders;
            border.LineStyle = XlLineStyle.xlContinuous;
            border.Weight = 2d;
            
        }

    }
}
