// Used ClosedXML(in order to read and modify the excel files) and tried using 
// FreeSpire.XLS (to convert from XLS to XLSX format in order to use ClosedXML but the free version allows for
// converting just the first 200 rows in each sheet so I converted manually outside the code the xls file to XLSX) 
// Used Microsoft.Office.Interop.Excel to convert from .xls to .xlsx format
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;

namespace ExcelFileModifier
{
    class ExcelModifier
    {
        static string excelFile;
        static string excelFileTask1;
        static string excelFileTask2;
        static int sheet1RowsCount = 0;
        static int sheet2RowsCount = 0;


        [STAThread]
        static void Main(string[] args)
        {
            char[] separator = { '.' };
            ExcelModifier efm = new ExcelModifier();
            efm.SetFilePath();
            if (excelFile != "") { 
            if (excelFile.Split(separator)[1].Equals("xls")) {
                efm.ConvertXLS_XLSX(excelFile);
                excelFile += "x";
            }
                if (excelFile.Split(separator)[1].Equals("xlsx"))
                {
                    try
                    {
                        excelFileTask1 = excelFile.Split(separator)[0] + "Task1.xlsx";
                        excelFileTask2 = excelFile.Split(separator)[0] + "Task2.xlsx";
                        var workbook = new XLWorkbook(excelFile);
                        var worksheet1 = workbook.Worksheet(1);
                        var worksheet2 = workbook.Worksheet(2);
                        efm.SetSheetRowsCount(workbook, worksheet1, worksheet2);
                        efm.SplitSheet1Adress(workbook, worksheet1);
                        workbook = new XLWorkbook(excelFile);
                        efm.SetWaxID(workbook);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }
                else Console.WriteLine("The provided file is not an Excel file");

            } else Console.WriteLine("No file was selected");
            
            Console.ReadKey();
        }

        public void SetFilePath()
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx; *.xls;"
            };
            ofd.ShowDialog();
            excelFile = ofd.FileName;
        }

        private void SetWaxID(XLWorkbook wb)
        {
            Console.WriteLine("Task 2 started");
            var ws1 = wb.Worksheet(1);
            var ws2 = wb.Worksheet(2);
            string rowData1 = "";
            string rowData2 = "";

            for (int i = 2; i < sheet2RowsCount; i++)
            {
                rowData2 = ws2.Cell("B" + i).Value.ToString() + ws2.Cell("C" + i).Value.ToString() + ws2.Cell("D" + i).Value.ToString();
                for (int j = 2; j < sheet1RowsCount; j++)
                {
                    rowData1 = ws1.Cell("B" + j).Value.ToString() + ws1.Cell("C" + j).Value.ToString() + ws1.Cell("D" + j).Value.ToString();
                    if (rowData1.Equals(rowData2))
                    {
                        ws2.Cell("A" + i).Value = ws1.Cell("A" + j).Value;
                        break;
                    }
                }
            }
            wb.SaveAs(excelFileTask2);
            Console.WriteLine("Task 2 completed");

        }

        private void SetSheetRowsCount(XLWorkbook wb, IXLWorksheet ws1, IXLWorksheet ws2)
        {
            sheet1RowsCount = ws1.RowsUsed().Count();
            sheet2RowsCount = ws2.RowsUsed().Count();
        }

        private void SplitSheet1Adress(XLWorkbook wb, IXLWorksheet ws1)
        {
            Console.WriteLine("Task 1 Started");
            string address = "";
            string town = "";
            char[] separator = { ',' };
            ws1.Cell("E1").Value = "Town";
            for (int i = 2; i <= sheet1RowsCount; i++)
            {
                if (ws1.Cell("D" + i).Value.ToString().Contains(','))
                {
                    address = ws1.Cell("D" + i).Value.ToString().Split(separator)[0];
                    town = ws1.Cell("D" + i).Value.ToString().Split(separator)[1];
                    ws1.Cell("D" + i).Value = address;
                    ws1.Cell("E" + i).Value = town;
                }
             
               
            }

            wb.SaveAs(excelFileTask1);
            Console.WriteLine("Task 1 Completed");

        }

       
        private void ConvertXLS_XLSX(string filePath)
        {
           
            var file = new FileInfo(filePath);
            var app = new Microsoft.Office.Interop.Excel.Application();
            var xlsFile = file.FullName;
            var wb = app.Workbooks.Open(xlsFile);
            var xlsxFile = xlsFile + "x";
            wb.SaveAs(Filename: xlsxFile, FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();
            Console.WriteLine("Converted from .xls to .xlsx format");
        }


        /* public void ConvertXLSToXLSX()
         {
             Workbook wb = new Workbook();
             wb.LoadFromFile(excelFile);
             wb.SaveToFile(excelFile + "x", ExcelVersion.Version2013);
             excelFile += "x";
         }
         */
    }
}
