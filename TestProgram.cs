
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace TestApp
{
    class TestProgram
    {
        public static void Main(string[] args)
        {
            // ToDo - we would need to put in a function to browse for the XLSX file
            // do some basic error checking to make sure it's an XLSX file
            // and also to browse the default/first sheet or ability to choose the sheet

            // string filepath = "C:\\Users\\Glen Millard\\Documents\\Second_receipts.xlsx";

            Console.Write("Enter the full path of your XLSX file: ");
            string filepath = Console.ReadLine();

            DataTable dt = ReadXLSX(filepath);

            // print headers:
            string headers = "";
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                headers += dt.Columns[i].ColumnName + " ";
            }
            
            Console.WriteLine(headers);

            foreach (DataRow row in dt.Rows)
            {
                string s_row = "";
                
                foreach (DataColumn column in dt.Columns)
                {
                    s_row += row[column].ToString() + "  ";
                }

                Console.WriteLine(s_row);
            }


        }

        public static DataTable ReadXLSX(string filepath)
        {
            FileInfo existingFile = new FileInfo(filepath);
            DataTable dt = new DataTable();

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {

                // Reads in first wprksheet of Excel file
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                if (worksheet.Dimension == null)
                {
                    return dt;
                }
                
                else 
                {

                    List<string> columnNames = new List<string>();
                    int currentColumn = 1;
                    foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                    {
                        string columnName = cell.Text.Trim();

                        //check if the previous header was empty and add it if it was
                        if (cell.Start.Column != currentColumn)
                        {
                            columnNames.Add("Header_" + currentColumn);
                            dt.Columns.Add("Header_" + currentColumn);
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
                        dt.Columns.Add(columnName);

                        currentColumn++;
                    }

                    //start adding the contents of the excel file to the datatable
                    for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
                    {
                        var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
                        DataRow newRow = dt.NewRow();

                        //loop all cells in the row
                        foreach (var cell in row)
                        {
                            newRow[cell.Start.Column - 1] = cell.Text;
                        }

                        dt.Rows.Add(newRow);
                    }

                    return dt;
                }

            }
        }

    }
}
