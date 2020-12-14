using System;
using System.IO;
using System.Data;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReader
{
    class ExcelHandler
    {
        public static ExcelPackage ExcelFileReader(string fileName)
        {
            if (fileName == null || fileName.Length == 0)
            {
                //return null;
                throw new ArgumentNullException("FileName is not exist.");
            }
            FileInfo fileInfo = new FileInfo(fileName);
            if (!fileInfo.Exists)
            {
                throw new FileNotFoundException("The file was not found.", fileName);
                //return null;
            }

            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            return excelPackage;
        }
        public static DataTable[] ExcelToDataTable(ExcelPackage excelPackage)
        {
            if (excelPackage == null)
            {
                throw new ArgumentNullException("The ExcelPackage was not found.");
                //Console.WriteLine("파일없음");
                //return null;
            }

            int workSheetsCounter = excelPackage.Workbook.Worksheets.Count;
            DataTable[] dataTable = new DataTable[workSheetsCounter];
            ExcelWorksheet[] worksheets = new ExcelWorksheet[workSheetsCounter];

            for (int i = 0; i < workSheetsCounter; i++)
            {
                worksheets[i] = excelPackage.Workbook.Worksheets[i];
                dataTable[i] = new DataTable();
            }

            //check if the worksheet is completely empty
            if (worksheets[0].Dimension == null)
            {
                return dataTable;
            }
            
            for (int i = 0; i < workSheetsCounter; i++)
            {
                //create a list to hold the column names
                List<string> columnNames = new List<string>();

                //needed to keep track of empty column headers
                int currentColumn = 1;
                //loop all columns in the sheet and add them to the datatable
                //헤더가 1행이 아닌경우 대응할것 - 미리 헤더 이외의 행 삭제해서 인풋하도록
                for (int j = 1; j <= worksheets[i].Dimension.End.Column; j++)
                {
                    string columnName = worksheets[i].Cells[1, j].Text.Trim();
                    if (columnName == null || columnName == "")
                        columnName = "Empty";

                    //check if the previous header was empty and add it if it was
                    if (worksheets[i].Cells[1, j].Start.Column != currentColumn)
                    {
                        columnNames.Add("Header_" + currentColumn);
                        dataTable[i].Columns.Add("Header_" + currentColumn);
                        currentColumn++;
                    }

                    //add the column name to the list to count the duplicates
                    columnNames.Add(columnName);

                    //count the duplicate column names and make them unique to avoid the exception
                    //A column named 'Name' already belongs to this DataTable
                    int occurrences = columnNames.Count(x => x.Equals(columnName));
                    if (occurrences > 1)
                    {
                        columnName = columnName + "_" + occurrences;
                    }

                    //add the column to the datatable
                    dataTable[i].Columns.Add(columnName);

                    currentColumn++;
                }
                
                //start adding the contents of the excel file to the datatable
                for (int k = 2; k <= worksheets[i].Dimension.End.Row; k++)
                {
                    var row = worksheets[i].Cells[k, 1, k, worksheets[i].Dimension.End.Column];
                    DataRow newRow = dataTable[i].NewRow();

                    //loop all cells in the row
                    foreach (var cell in row)
                    {
                        newRow[cell.Start.Column - 1] = cell.Text;
                    }

                    dataTable[i].Rows.Add(newRow);
                }
            }
            return dataTable;
        }
        public static void ShowTable(DataTable[] dataTable)
        {
            if (dataTable == null)
            {
                throw new ArgumentNullException("The DataTable was not found.");
                //return;
            }
            for (int i = 0; i < dataTable.Length; i++)
            {
                foreach (DataColumn col in dataTable[i].Columns)
                {
                    Console.Write("{0,-14}", col.ColumnName);
                }
                Console.WriteLine();

                foreach (DataRow row in dataTable[i].Rows)
                {
                    foreach (DataColumn col in dataTable[i].Columns)
                    {
                        if (col.DataType.Equals(typeof(DateTime)))
                            Console.Write("{0,-14:d}", row[col]);
                        else if (col.DataType.Equals(typeof(Decimal)))
                            Console.Write("{0,-14:C}", row[col]);
                        else
                            Console.Write("{0,-14}", row[col]);
                    }
                    Console.WriteLine();
                }
                Console.WriteLine();
            }
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string fileName = @"D:\honorscs\DataAnalyzing\ExcelReader\bin\Debug\netcoreapp3.1\a.xlsx";
            ExcelPackage excelfile = ExcelHandler.ExcelFileReader(fileName);
            DataTable[] excelDataTable = ExcelHandler.ExcelToDataTable(excelfile);

            ExcelHandler.ShowTable(excelDataTable);


        }
    }
}
