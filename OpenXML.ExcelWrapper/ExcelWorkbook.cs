using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace OpenXML.ExcelWrapper
{
    public class ExcelWorkbook
    {
        private readonly List<ExcelSheet> sheets = new List<ExcelSheet>();

        public ExcelWorkbook()
        {
        }

        public void AddSheet(ExcelSheet sheet)
        {
            this.sheets.Add(sheet);
        }

        public byte[] Save()
        {
            using (var ms = new MemoryStream())
            using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                Sheets sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());

                var stylesPart = doc.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new ExcelWorkbookStylesheet().CreateStylesheet(this.sheets);

                foreach (var item in this.sheets)
                {
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var workSheet = new Worksheet();
                    var sheetData = new SheetData();

                    foreach (var rowGroup in item.Cells.GroupBy(g => g.Value.Row))
                    {
                        var row = new Row() { RowIndex = Convert.ToUInt32(rowGroup.Key) };
                        foreach (var cellItem in rowGroup)
                            row.Append(this.CreateCell(cellItem.Value));

                        sheetData.Append(row);
                    }

                    workSheet.AppendChild(sheetData);
                    worksheetPart.Worksheet = workSheet;

                    uint sheetId = 1;
                    if (sheets.Elements<Sheet>().Count() > 0)
                        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;

                    Sheet sheet1 = new Sheet()
                    {
                        Id = doc.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = sheetId,
                        Name = item.SheetName
                    };

                    sheets.Append(sheet1);
                }

                doc.Close();
                return ms.ToArray();
            }
        }

        public static ExcelWorkbook FromFile(string fileName)
        {
            throw new NotImplementedException();
        }

        public static ExcelWorkbook FromStream(Stream stream)
        {
            throw new NotImplementedException();
        }

        private Cell CreateCell(ExcelCell excelCell)
        {
            if (excelCell is null)
                throw new ArgumentNullException(nameof(excelCell));

            if (excelCell.Value is null)
            {
                var emptyCell = new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(""),
                };

                return emptyCell;
            }

            if (excelCell.IsFormula)
            {
                var cellWithFormula = new Cell
                {
                    CellFormula = new CellFormula(excelCell.Value as string),
                    CellValue = new CellValue()
                };

                return cellWithFormula;
            }

            CellValue cellValue = null;
            CellValues excelCellType = CellValues.String;

            if (excelCell.ValueType == typeof(decimal)
                || excelCell.ValueType == typeof(decimal?)
                || excelCell.ValueType == typeof(int)
                || excelCell.ValueType == typeof(int?)
                || excelCell.ValueType == typeof(float)
                || excelCell.ValueType == typeof(float?)
                || excelCell.ValueType == typeof(double)
                || excelCell.ValueType == typeof(double?))
            {
                excelCellType = CellValues.Number;
                cellValue = new CellValue(string.Format(CultureInfo.InvariantCulture, "{0:N8}", excelCell.Value));
            }
            else if (excelCell.ValueType == typeof(DateTime)
                || excelCell.ValueType == typeof(DateTime?))
            {
                //https://stackoverflow.com/questions/2792304/how-to-insert-a-date-to-an-open-xml-worksheet
                excelCellType = CellValues.Number;
                cellValue = new CellValue(((DateTime)excelCell.Value).ToOADate().ToString(CultureInfo.InvariantCulture));
            }
            else if (excelCell.ValueType == typeof(bool)
                || excelCell.ValueType == typeof(bool?))
            {
                excelCellType = CellValues.Number;
                cellValue = new CellValue((bool?)excelCell.Value == true ? "1" : "0");
            }
            else
            {
                excelCellType = CellValues.String;
                cellValue = new CellValue(excelCell.Value.ToString());
            }

            var cell = new Cell
            {
                DataType = excelCellType,
                CellValue = cellValue,
            };

            if (excelCell.CellStyle != null)
                cell.StyleIndex = excelCell.CellStyle.StyleIndex;

            return cell;
        }
    }
}
