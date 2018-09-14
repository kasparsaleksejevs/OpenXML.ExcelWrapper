using OpenXML.ExcelWrapper;
using System;
using System.IO;

namespace ExcelWrapperConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var wb = new ExcelWorkbook();
            var myFirstSheet = new ExcelSheet("My Sheet");
            wb.AddSheet(myFirstSheet);




            myFirstSheet.AddOrUpdateCell(new ExcelCell("A3", "Text 1"));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("B3", "Other text"));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("C3", "C Column"));
            myFirstSheet.AddOrUpdateCell(new ExcelCell(4, 3, "D Column"));

            myFirstSheet.AddOrUpdateCell(new ExcelCell("A", 4, 0.34m) { CellFormat = CellFormatEnum.DecimalTwoDecimals });
            myFirstSheet.AddOrUpdateCell(new ExcelCell("B", 4, 0.231) { CellFormat = CellFormatEnum.PercentageTwoDecimals });
            myFirstSheet.AddOrUpdateCell(new ExcelCell("C", 4, DateTime.Now));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("D", 4, "ZZZ"));

            myFirstSheet.AddOrUpdateCell(new ExcelCell("A", 5, 0.10m));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("B", 5, 0.20m));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("C", 5, 0.50m));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("D", 5, 0.99m));

            //myFirstSheet.AddFormula("A", 6, "SUM(A5:D5)");

            var xlsData = wb.Save();

            var fileName = @"C:\temp\MyExcel.xlsx";

            File.WriteAllBytes(fileName, xlsData);
            System.Diagnostics.Process.Start(fileName);
        }
    }
}
