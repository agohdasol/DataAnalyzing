using System;
using System.IO;
using System.Data;
using OfficeOpenXml;
using System.Collections.Generic;

namespace ExcelReader
{
    class ExcelHandler
    {
        public static ExcelPackage ExcelFileReader(string fileName)
        {
            FileInfo fileInfo = new FileInfo(fileName);
            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            return excelPackage;
        }
        public static DataTable[] ExcelToDataTable(ExcelPackage excelPackage)
        {
            int workSheetsCounter = excelPackage.Workbook.Worksheets.Count;
            DataTable[] dataTable = new DataTable[workSheetsCounter];
            ExcelWorksheet[] worksheets = new ExcelWorksheet[workSheetsCounter];
            worksheets[0] = excelPackage.Workbook.Worksheets[1];
            for (int i = 0; i < workSheetsCounter; i++)
            {
                try
                {
                    worksheets[i] = excelPackage.Workbook.Worksheets[i];
                }
                catch (IndexOutOfRangeException e)
                {
                    workSheetsCounter += 1; //워크시트 인덱스가 1부터 시작하는 경우에는 인덱스 0인 경우 패스하고 시트 갯수 값을 1 증가시킴
                    continue;
                }
            }



            ////check if the worksheet is completely empty
            //if (worksheet.Dimension == null)
            //{
            //    return dataTable;
            //}

            ////create a list to hold the column names
            //List<string> columnNames = new List<string>();

            ////needed to keep track of empty column headers
            //int currentColumn = 1;

            ////loop all columns in the sheet and add them to the datatable
            //foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            //{
            //    string columnName = cell.Text.Trim();

            //    //check if the previous header was empty and add it if it was
            //    if (cell.Start.Column != currentColumn)
            //    {
            //        columnNames.Add("Header_" + currentColumn);
            //        dataTable.Columns.Add("Header_" + currentColumn);
            //        currentColumn++;
            //    }

            //    //add the column name to the list to count the duplicates
            //    columnNames.Add(columnName);

            //    //count the duplicate column names and make them unique to avoid the exception
            //    //A column named 'Name' already belongs to this DataTable
            //    //int occurrences = columnNames.Count(x => x.Equals(columnName));
            //    //if (occurrences > 1)
            //    //{
            //    //    columnName = columnName + "_" + occurrences;
            //    //}

            //    //add the column to the datatable
            //    dataTable.Columns.Add(columnName);

            //    currentColumn++;
            //}

            ////start adding the contents of the excel file to the datatable
            //for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
            //{
            //    var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
            //    DataRow newRow = dataTable.NewRow();

            //    //loop all cells in the row
            //    foreach (var cell in row)
            //    {
            //        newRow[cell.Start.Column - 1] = cell.Text;
            //    }

            //    dataTable.Rows.Add(newRow);
            //}

            return dataTable[];
        }
        public static void ShowTable(DataTable dataTable)
        {
            foreach(DataColumn col in dataTable.Columns)
            {
                Console.Write("{0,-14}", col.ColumnName);
            }
            Console.WriteLine();

            foreach(DataRow row in dataTable.Rows)
            {
                foreach(DataColumn col in dataTable.Columns)
                {
                    if (col.DataType.Equals(typeof(DateTime)))
                        Console.Write("{0,-14:d}",row[col]);
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
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string fileName = "C:\\cs\\ExcelReader\\bin\\Debug\\netcoreapp3.1\\a.xlsx";
            ExcelPackage excelfile = ExcelHandler.ExcelFileReader(fileName);
            DataTable excelDataTable = ExcelHandler.ExcelToDataTable(excelfile);

            ExcelHandler.ShowTable(excelDataTable);


        }
    }
}
