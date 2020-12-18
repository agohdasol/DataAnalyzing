using System;
using System.IO;
using System.Data;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;
using System.Collections;
using System.Text.RegularExpressions;

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
        //시트명으로 불러오는 메서드도 구현?-시트명 입력받고 시트명 찾아서 인덱스 반환->메소드 오버로드
        public static DataTable ExcelToDataTable(ExcelPackage excelPackage, int sheetIndex = 0, int headerIndex = 1, int skipFooter = 4)
        {
            if (excelPackage == null)
            {
                throw new ArgumentNullException("The ExcelPackage was not found.");
                //return null;
            }

            int workSheetsCounter = excelPackage.Workbook.Worksheets.Count;
            DataTable dataTable = new DataTable();
            if (workSheetsCounter < sheetIndex + 1)
                //throw new IndexOutOfRangeException("The SheetIndex is out of index.");
                return dataTable;
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[sheetIndex];

            //check if the worksheet is completely empty
            if (worksheet.Dimension == null)
            {
                return dataTable;
            }

            //create a list to hold the column names
            List<string> columnNames = new List<string>();

            //needed to keep track of empty column headers
            int currentColumn = 1;
            //loop all columns in the sheet and add them to the datatable
            if (headerIndex >= worksheet.Dimension.End.Row)
                //return dataTable;
                throw new IndexOutOfRangeException("The input header is out of index.");

            //Get Type of each column
            Type[] typeArray = ColumnTypeParser(worksheet, headerIndex, skipFooter);
            for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
            {
                string columnName = worksheet.Cells[1, i].Text.Trim();
                if (columnName == null || columnName == "")
                    columnName = "Empty";

                //check if the previous header was empty and add it if it was
                if (worksheet.Cells[1, i].Start.Column != currentColumn)
                {
                    columnNames.Add("Header_" + currentColumn);
                    dataTable.Columns.Add("Header_" + currentColumn, typeArray[i - 1]);
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
                dataTable.Columns.Add(columnName, typeArray[i - 1]);

                currentColumn++;
            }

            //start adding the contents of the excel file to the datatable
            if (headerIndex + 1 > worksheet.Dimension.End.Row - skipFooter)
                //return dataTable;
                throw new IndexOutOfRangeException("The SkipFooter is out of index.");
            for (int i = headerIndex + 1; i <= worksheet.Dimension.End.Row - skipFooter; i++)
            {
                var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
                DataRow newRow = dataTable.NewRow();

                //loop all cells in the row
                foreach (var cell in row)
                {
                    if (decimal.TryParse(cell.Value.ToString(), out decimal temporaryDecimal))
                        newRow[cell.Start.Column - 1] = temporaryDecimal;
                    else if (int.TryParse(cell.Value.ToString(), out int temporaryInt))
                        newRow[cell.Start.Column - 1] = temporaryInt;
                    else
                        newRow[cell.Start.Column - 1] = cell.Text;
                }

                dataTable.Rows.Add(newRow);
            }
            return dataTable;
        }
        private static Type[] ColumnTypeParser(ExcelWorksheet worksheet, int headerIndex = 1, int skipFooter = 0)
        {
            Type[] typeArray = new Type[worksheet.Dimension.End.Column];
            //헤더인덱스 크기는 ExcelToDataTable메서드에서 검증하므로 별도 검증 생략
            for (int i = headerIndex + 1; i <= worksheet.Dimension.End.Row - skipFooter; i++)
            {
                //TryParse out에 이용될 변수 선언
                //int temporaryInt;
                //decimal temporaryDecimal;
                //loop all cells in the row
                for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                {
                    //정수-데시멀-스트링 순으로 판단. 날짜 등은 인식X
                    if (typeArray[j - 1] == null)
                    {
                        if (int.TryParse(worksheet.Cells[i, j].Value.ToString(), out int _))
                        {
                            typeArray[j - 1] = typeof(int);
                        }
                        else if (decimal.TryParse(worksheet.Cells[i, j].Value.ToString(), out decimal _))
                            typeArray[j - 1] = typeof(decimal);
                        else
                            typeArray[j - 1] = typeof(string);
                    }
                    else
                    {
                        //기존 입력값이 인트인데 새로운 입력값이 데시멀인 경우 해당 칼럼은 데시멀
                        if (typeArray[j - 1] == typeof(int)
                            && !int.TryParse(worksheet.Cells[i, j].Value.ToString(), out int _)
                            && decimal.TryParse(worksheet.Cells[i, j].Value.ToString(), out decimal _))
                            typeArray[j - 1] = typeof(decimal);
                        //기존 입력값이 데시멀인데 새로운 입력값이 데시멀도 인트도 아닌경우 해당 칼럼은 스트링
                        else if (typeArray[j - 1] == typeof(decimal)
                            && !decimal.TryParse(worksheet.Cells[i, j].Value.ToString(), out _)
                            && !int.TryParse(worksheet.Cells[i, j].Value.ToString(), out _))
                            typeArray[j - 1] = typeof(string);
                    }

                }
            }
            return typeArray;
        }
        public static void ShowTable(DataTable dataTable)
        {
            if (dataTable == null)
            {
                throw new ArgumentNullException("The DataTable was not found.");
                //return;
            }
            foreach (DataColumn col in dataTable.Columns)
            {
                Console.Write("{0,-10}", col.ColumnName);
            }
            Console.WriteLine();

            foreach (DataRow row in dataTable.Rows)
            {
                foreach (DataColumn col in dataTable.Columns)
                {
                    if (col.DataType.Equals(typeof(DateTime)))
                        Console.Write("{0,-10:d}", row[col]);
                    //else if (col.DataType.Equals(typeof(Decimal)))
                    //    Console.Write("{0,-10}", row[col]);
                    else
                        Console.Write("{0,-10}", row[col]);
                }
                Console.WriteLine();
            }
            Console.WriteLine();
        }
        public static string ReadYearMonthFromHeader(ExcelPackage excelPackage) //챕터 4 전용
        {
            if (excelPackage == null)
            {
                throw new ArgumentNullException("The ExcelPackage was not found.");
                //return null;
            }

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];

            //check if the worksheet is completely empty
            if (worksheet.Dimension == null)
            {
                throw new ArgumentNullException("The Worksheet 0 is empty.");
            }

            string columnName = worksheet.Cells[1, 1].Text.Trim();
            //2019년 04월 외래객 입국-목적별/국적별
            Regex regex = new Regex(@"(?<Year>\d+)년 (?<Month>\d+)월 외래객 입국-목적별/국적별");

            if (regex.IsMatch(columnName))
                return regex.Match(columnName).Groups["Year"].Value.ToString() + "-"
                    + regex.Match(columnName).Groups["Month"].Value.ToString();
            else
                return "";
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string fileName = @"D:\honorscs\DataAnalyzing\ExcelReader\bin\Debug\netcoreapp3.1\filetest.xlsx";
            ExcelPackage excelfile = ExcelHandler.ExcelFileReader(fileName);
            DataTable excelDataTable = ExcelHandler.ExcelToDataTable(excelfile, headerIndex: 1, skipFooter: 4);

            ExcelHandler.ShowTable(excelDataTable);

            Console.WriteLine(excelDataTable.Rows[2]["성장률(%)"].GetType());
            Console.WriteLine(excelDataTable.Rows[2]["관광"].GetType());
            Console.WriteLine(excelDataTable.Rows[2]["성장률(%)"].ToString());
            Console.WriteLine(excelDataTable.Rows[2]["국적"].ToString());

            DataTable filtered = new DataTable();
            decimal test = 10.0m;
            filtered = excelDataTable.AsEnumerable()
                .Where(Row => Row.Field<decimal>("성장률(%)") > test)
                .CopyToDataTable();

            ExcelHandler.ShowTable(filtered);

            string fileName2 = @"D:\honorscs\DataAnalyzing\ExcelReader\bin\Debug\netcoreapp3.1\test22.xlsx";
            ExcelPackage excelfile2 = ExcelHandler.ExcelFileReader(fileName2);
            Console.WriteLine(ExcelHandler.ReadYearMonthFromHeader(excelfile2));
            //1행 읽어서 기준연월 체크하는 메서드 작성할것
            //여러파일읽어서 합치는메서드
        }
    }
}
