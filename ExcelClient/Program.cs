using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace ExcelClient
{
    class Program
    {
        static void Main(string[] args)
        {
            string fullPath = "C:\\Users\\TUNC\\source\\repos\\ExcelClient\\ExcelClient\\ExampleData\\kvkk.xlsx";
            using (var stream = File.Open(fullPath, FileMode.Open, FileAccess.Read))
            {
                Reader reader = new Reader(stream);
                //reader.GetRow(0, "Envanter örneği");
                //reader.GetColumn(1, "Envanter örneği");
                //reader.GetTable("Envanter örneği");

            }

            DataTableManager dataTableManager = new DataTableManager();
            dataTableManager.DataTable.Columns.Add("Selam1",typeof(string));
            dataTableManager.DataTable.Columns.Add("Selam2", typeof(string));
            dataTableManager.DataTable.Columns.Add("Selam3", typeof(int));
            dataTableManager.DataTable.Rows.Add("asxasxs","xasx",1);
            dataTableManager.DataTable.Rows.Add("sdasdas","1swqs",1);
            dataTableManager.DataTable.Rows.Add("asxadsasasassxs","sxasxa",5);
            
            Writer writer = new Writer(dataTableManager.DataTable, "SheetNameExample","Table 1", @"C:\Book1.xlsx");
            writer.CloseExcel();
            
            Console.ReadKey();
        }
    }
}
