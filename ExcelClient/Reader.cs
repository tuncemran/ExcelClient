using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;

namespace ExcelClient
{
    class Reader
    {
        private FileStream FileStream{ get; set; }
        public Reader(FileStream stream)
        {
            FileStream = stream;
        }

        public List<object> GetColumn(int columnId,string sheetName)
        {
            List<object> columnList = new List<object>();
            try
            {
                using (var reader = ExcelReaderFactory.CreateReader(FileStream))
                {
                    var conf = new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = a => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        }
                    };
                    var dataSet = reader.AsDataSet(conf);
                    var sheet = dataSet.Tables[sheetName].Rows.Cast<DataRow>(); // instead of sheetName you can use Index of it like 0 ,1 , ...
                    foreach (var row in sheet)
                    {
                        var rowId = row.Table.Rows.IndexOf(row);
                        var rowValueByHeaderFieldName = row[columnId]; 
                        columnList.Add(rowValueByHeaderFieldName);
                    }
                }

            }
            catch (Exception e)
            {
                
            }

            return columnList;
        }

        public List<object> GetRow(int rowId, string sheetName)
        {
            List<object> row = new List<object>();
            
            try
            {
                using (var reader = ExcelReaderFactory.CreateReader(FileStream))
                {
                    var conf = new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = a => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        }
                    };
                    var dataSet = reader.AsDataSet(conf);
                    var sheet = dataSet.Tables[sheetName].Rows.Cast<DataRow>(); // instead of sheetName you can use Index of it like 0 ,1 , ...
                    int count = 0;
                    foreach (var rowObject in sheet)
                    {
                        if (count == rowId)
                        {
                            var rowItem = rowObject.ItemArray;
                            row = rowItem.OfType<object>().ToList();
                            break;
                        }
                        else
                        {
                            count++;
                        }
                        

                    }
                }

            }
            catch (Exception e)
            {

            }

            return row;
        }

        public List<object> GetTable(string sheetName)
        {

            List<object> table = new List<object>();
            try
            {
                using (var reader = ExcelReaderFactory.CreateReader(FileStream))
                {
                    var conf = new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = a => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        }
                    };
                    var dataSet = reader.AsDataSet(conf);
                    var sheet = dataSet.Tables[sheetName].Rows.Cast<DataRow>(); // instead of sheetName you can use Index of it like 0 ,1 , ...
                    
                    foreach (var rowObject in sheet)
                    {
                        var rowItem = rowObject.ItemArray;
                        table.Add(rowItem.OfType<object>().ToList());
                    }
                }

            }
            catch (Exception e)
            {

            }

            return table;
        }
    }
}
