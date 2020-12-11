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

            int workSheetsCounter = excelPackage.Workbook.Worksheets.Count; //워크시트 인덱스와 데이터테이블 인덱스가 다른경우?
            DataTable[] dataTable = new DataTable[workSheetsCounter];
            ExcelWorksheet[] worksheets = new ExcelWorksheet[workSheetsCounter];

            bool is0BaseWorkSheets = true;  //인덱스가 0부터 시작하는 워크시트인지 여부
            for (int i = 0; i < workSheetsCounter; i++)
            {
                try
                {
                    worksheets[i] = excelPackage.Workbook.Worksheets[i];
                }
                catch (IndexOutOfRangeException)
                {
                    workSheetsCounter += 1; //워크시트 인덱스가 1부터 시작하는 경우에는 인덱스 0인 경우 패스하고 시트 갯수 값을 1 증가시킴
                    is0BaseWorkSheets = false;
                    continue;
                }
            }
            //데이터테이블 초기화할것
            //check if the worksheet is completely empty
            if ((is0BaseWorkSheets? worksheets[0].Dimension : worksheets[1].Dimension) == null)
            {
                return dataTable;
            }

            //create a list to hold the column names
            List<string> columnNames = new List<string>();

            //needed to keep track of empty column headers
            int currentColumn = 1;

            for(int i = (is0BaseWorkSheets? 0:1); i < workSheetsCounter; i++)
            {
                //loop all columns in the sheet and add them to the datatable
                foreach (var cell in worksheets[i].Cells[1, 1, 1, worksheets[i].Dimension.End.Column])
                {
                    string columnName = cell.Text.Trim();

                    //check if the previous header was empty and add it if it was
                    if (cell.Start.Column != currentColumn)
                    {
                        columnNames.Add("Header_" + currentColumn);
                        dataTable[i].Columns.Add("Header_" + currentColumn);
                        currentColumn++;
                    }

                    //add the column name to the list to count the duplicates
                    columnNames.Add(columnName);

                    //count the duplicate column names and make them unique to avoid the exception
                    //A column named 'Name' already belongs to this DataTable
                    //int occurrences = columnNames.Count(x => x.Equals(columnName));
                    //if (occurrences > 1)
                    //{
                    //    columnName = columnName + "_" + occurrences;
                    //}

                    //add the column to the datatable
                    dataTable[i].Columns.Add(columnName);

                    currentColumn++;
                }

                //start adding the contents of the excel file to the datatable
                for (int j = 2; j <= worksheets[i].Dimension.End.Row; j++)
                {
                    var row = worksheets[i].Cells[j, 1, j, worksheets[i].Dimension.End.Column];
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
            for (int i=0;i<dataTable.Length;i++)
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
