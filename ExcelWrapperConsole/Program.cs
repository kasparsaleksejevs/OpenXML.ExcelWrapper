using OpenXML.ExcelWrapper;
using System;
using System.IO;

namespace ExcelWrapperConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var wb = ExcelCreatorWorkbook.CreateWorkbook();
            var myFirstSheet = wb.AddSheet("My Sheet");

            myFirstSheet.AddCell("A", 3, "Text 1");
            myFirstSheet.AddCell("B", 3, "Other text");
            myFirstSheet.AddCell("C", 3, "C Column");
            myFirstSheet.AddCell(ExcelCell.GetColumnLetters(4), 3, "D Column");

            myFirstSheet.AddCell("A", 4, 0.34m, CellFormatEnum.DecimalTwoDecimals);
            myFirstSheet.AddCell("B", 4, 0.231, CellFormatEnum.PercentageTwoDecimals);
            myFirstSheet.AddCell("C", 4, DateTime.Now);
            myFirstSheet.AddCell("D", 4, "ZZZ");

            myFirstSheet.AddCell("A", 5, 0.10m);
            myFirstSheet.AddCell("B", 5, 0.20m);
            myFirstSheet.AddCell("C", 5, 0.50m);
            myFirstSheet.AddCell("D", 5, 0.99m);

            myFirstSheet.AddFormula("A", 6, "SUM(A5:D5)");

            var xlsData = wb.Save();

            var fileName = @"C:\temp\MyExcel.xlsx";

            File.WriteAllBytes(fileName, xlsData);
            System.Diagnostics.Process.Start(fileName);
        }
    }
}
