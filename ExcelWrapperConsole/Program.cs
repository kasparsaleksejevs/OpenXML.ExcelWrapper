using OpenXML.ExcelWrapper;
using System;
using System.Drawing;
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

            var borderedYellowCell = new ExcelCellStyle
            {
                CellFormat = CellFormatEnum.PercentageTwoDecimals,
                BackgroundColor = Color.Yellow,
                FontColor = Color.Red
            };

            var greenCell = new ExcelCellStyle
            {
                BackgroundColor = Color.Green,
            };

            var bordersCell = new ExcelCellStyle
            {
                Borders = ExcelCellBorderEnum.Top | ExcelCellBorderEnum.Bottom
            };

            myFirstSheet.AddOrUpdateCell(new ExcelCell("A3", "Decimals"));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("B3", "Percentages"));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("C3", "C Column"));
            myFirstSheet.AddOrUpdateCell(new ExcelCell(4, 3, "D Column"));

            myFirstSheet.AddOrUpdateCell(new ExcelCell("A", 4, 0.34m, CellFormatEnum.DecimalTwoDecimals));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("B", 4, 0.231, CellFormatEnum.PercentageTwoDecimals));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("C", 4, DateTime.Now, CellFormatEnum.DateTime));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("D", 4, 0.55m) { CellStyle = borderedYellowCell });

            myFirstSheet.AddOrUpdateCell(new ExcelCell("A", 5, 0.10m, CellFormatEnum.DecimalTwoDecimals));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("B", 5, 0.20m, CellFormatEnum.PercentageTwoDecimals));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("C", 5, DateTime.Now, CellFormatEnum.DateTime));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("D", 5, 0.99m));

            myFirstSheet.AddOrUpdateCell(new ExcelCell("A6", 30));
            myFirstSheet.AddOrUpdateCell(new ExcelCell("B6", 20) { CellStyle = borderedYellowCell });
            myFirstSheet.AddOrUpdateCell(new ExcelCell("C6", 10) { CellStyle = greenCell });
            myFirstSheet.AddOrUpdateCell(new ExcelCell("D6", 55));

            myFirstSheet.AddOrUpdateCell(new ExcelCell("C8", "=SUM(A6:D6)") { CellStyle = bordersCell });

            var xlsData = wb.Save();

            var fileName = @"C:\temp\MyExcel_v2.xlsx";

            File.WriteAllBytes(fileName, xlsData);
            System.Diagnostics.Process.Start(fileName);
        }
    }
}
