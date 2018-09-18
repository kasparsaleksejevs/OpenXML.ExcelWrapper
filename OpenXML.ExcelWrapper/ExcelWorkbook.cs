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
        /// <summary>
        /// The sheets to be written to the Excel document.
        /// </summary>
        private readonly List<ExcelSheet> sheets = new List<ExcelSheet>();

        private readonly string fileName = null;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWorkbook"/> class.
        /// </summary>
        public ExcelWorkbook()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWorkbook"/> class.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        public ExcelWorkbook(string fileName)
        {
            this.fileName = fileName;
        }

        public void AddSheet(ExcelSheet sheet)
        {
            this.sheets.Add(sheet);
        }

        public ExcelSheet GetSheetByName(string sheetName)
        {
            var sheet = new ExcelSheet(sheetName);
            this.sheets.Add(sheet);
            return sheet;
        }

        /// <summary>
        /// Saves the spreadsheet as a new file.
        /// </summary>
        /// <returns>Byte array containing spreadsheet data.</returns>
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

                    // ToDo: drawings part will go here

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

        /// <summary>
        /// Saves the spreadsheet to the existing file, adding or updating cells as specified.
        /// Note: only cell data types and values are updated. The cell style updates currently are not supported. Adding new sheets is also not yet supported.
        /// </summary>
        /// <returns>Byte array containing spreadsheet data.</returns>
        public void Update()
        {
            if (string.IsNullOrEmpty(this.fileName))
                throw new Exception("FileName not specified. Please instantiate ExcelWorkbook with fileName overload");

            using (var ms = new FileStream(this.fileName, FileMode.Open))
            using (var doc = SpreadsheetDocument.Open(ms, true))
            {
                var workbookPart = doc.WorkbookPart;
                var sheets = doc.WorkbookPart.Workbook.Sheets;

                foreach (var item in this.sheets)
                {
                    var sheet = sheets.OfType<Sheet>().FirstOrDefault(m => m.Name == item.SheetName);
                    if (sheet == null)
                        throw new Exception($"Sheet with name '{item.SheetName}' does not exist.");

                    var worksheetPart = doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    var worksheet = worksheetPart.Worksheet;

                    foreach (var rowGroup in item.Cells.GroupBy(g => g.Value.Row))
                    {
                        var rowIndex = (uint)rowGroup.Key;
                        var row = worksheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
                        if (row is null)
                        {
                            row = new Row() { RowIndex = rowIndex };
                            worksheet.Append(row);
                        }

                        foreach (var cellItem in rowGroup)
                        {
                            var cell = row.Elements<Cell>().FirstOrDefault(m => m.CellReference.Value == cellItem.Value.Address);
                            if (cell is null)
                            {
                                cell = this.CreateCell(cellItem.Value, withoutStyle: true);

                                var nextCell = row.Elements<Cell>().FirstOrDefault(m => m.CellReference.Value.CompareTo(cellItem.Value.Address) > 0);
                                if (nextCell != null)
                                    row.InsertBefore(cell, nextCell);
                                else
                                    row.AppendChild(cell);
                            }

                            cell.CellValue = new CellValue(cellItem.Value.Value.ToString());
                        }
                    }

                    worksheetPart.Worksheet.Save();
                }

                doc.Close();
            }
        }

        /// <summary>
        /// Loads the spreadsheet from file for reading (and writing).
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <returns>ExcelWorkbook instance.</returns>
        /// <exception cref="NotImplementedException">not yet implemented</exception>
        public static ExcelWorkbook FromFile(string fileName)
        {
            return new ExcelWorkbook(fileName);

        }

        /// <summary>
        /// Loads the spreadsheet from stream for reading (and writing).
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>ExcelWorkbook instance.</returns>
        /// <exception cref="NotImplementedException">not yet implemented</exception>
        public static ExcelWorkbook FromStream(Stream stream)
        {
            throw new NotImplementedException();
        }

        private Cell CreateCell(ExcelCell excelCell, bool withoutStyle = false)
        {
            if (excelCell is null)
                throw new ArgumentNullException(nameof(excelCell));

            var cell = new Cell { CellReference = excelCell.Address };
            if (excelCell.CellStyle != null && !withoutStyle)
                cell.StyleIndex = excelCell.CellStyle.StyleIndex;

            if (excelCell.Value is null)
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue("");
                return cell;
            }

            if (excelCell.IsFormula)
            {
                cell.CellFormula = new CellFormula(excelCell.Value as string);
                cell.CellValue = new CellValue();
                return cell;
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

            cell.DataType = excelCellType;
            cell.CellValue = cellValue;
            cell.CellReference = excelCell.Address;

            return cell;
        }
    }
}
