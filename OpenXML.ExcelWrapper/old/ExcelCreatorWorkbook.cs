using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;


namespace OpenXML.ExcelWrapper.old
{
    public class ExcelCreatorWorkbook
    {
        private readonly List<ExcelCreatorSheet> sheets = new List<ExcelCreatorSheet>();

        private ExcelCreatorWorkbook()
        {
        }



        public ExcelCreatorSheet AddSheet(string name)
        {
            var sheet = new ExcelCreatorSheet(name);
            this.sheets.Add(sheet);
            return sheet;
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
                stylesPart.Stylesheet = new Stylesheet
                {
                    Fonts = new Fonts(new Font()),
                    Fills = new Fills(new Fill()),
                    Borders = new Borders(new Border()),
                    CellStyleFormats = new CellStyleFormats(new CellFormat()),
                    CellFormats =
                        new CellFormats(
                            new CellFormat(),
                            new CellFormat
                            {
                                NumberFormatId = (uint)NumberFormats.DateTime,
                                ApplyNumberFormat = true
                            },
                            new CellFormat
                            {
                                NumberFormatId = (uint)NumberFormats.DecimalTwoDecimals,
                                ApplyNumberFormat = true
                            },
                            new CellFormat
                            {
                                NumberFormatId = (uint)NumberFormats.Percentage,
                                ApplyNumberFormat = true
                            },
                            new CellFormat
                            {
                                NumberFormatId = (uint)NumberFormats.PercentageTwoDecimals,
                                ApplyNumberFormat = true
                            })
                };

                foreach (var item in this.sheets)
                {
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var workSheet = new Worksheet();
                    var sheetData = new SheetData();

                    foreach (var rowGroup in item.Cells.GroupBy(g => g.Value.Row))
                    {
                        var row = new Row() { RowIndex = Convert.ToUInt32(rowGroup.Key) };
                        foreach (var cellItem in rowGroup)
                            row.Append(this.CreateCell(cellItem.Value.Value, cellItem.Value.CellFormatId));

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

        private Cell CreateCell(object dataCell, uint? styleIndex)
        {
            if (dataCell is null)
            {
                var emptyCell = new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(""),
                };

                return emptyCell;
            }

            if (dataCell.GetType() == typeof(ExcelCellFormula))
            {
                // https://social.msdn.microsoft.com/Forums/office/en-US/b28a9f9c-9ad7-4eea-adab-b062883c11e2/formula-cells-in-excel-using-openxml?forum=oxmlsdk
                var cellWithFormula = new Cell
                {
                    CellFormula = new CellFormula(((ExcelCellFormula)dataCell).Formula),
                    CellValue = new CellValue()
                };

                return cellWithFormula;
            }

            CellValue cellValue = null;
            CellValues excelCellType = CellValues.String;

            if (dataCell.GetType() == typeof(decimal)
                || dataCell.GetType() == typeof(decimal?)
                || dataCell.GetType() == typeof(int)
                || dataCell.GetType() == typeof(int?)
                || dataCell.GetType() == typeof(float)
                || dataCell.GetType() == typeof(float?)
                || dataCell.GetType() == typeof(double)
                || dataCell.GetType() == typeof(double?))
            {
                excelCellType = CellValues.Number;
                cellValue = new CellValue(string.Format(CultureInfo.InvariantCulture, "{0:N8}", dataCell));
            }
            else if (dataCell.GetType() == typeof(DateTime)
                || dataCell.GetType() == typeof(DateTime?))
            {
                //https://stackoverflow.com/questions/2792304/how-to-insert-a-date-to-an-open-xml-worksheet
                excelCellType = CellValues.Number;
                cellValue = new CellValue(((DateTime)dataCell).ToOADate().ToString(CultureInfo.InvariantCulture));
            }
            else if (dataCell.GetType() == typeof(bool)
                || dataCell.GetType() == typeof(bool?))
            {
                excelCellType = CellValues.Number;
                cellValue = new CellValue((bool?)dataCell == true ? "1" : "0");
            }
            else
            {
                excelCellType = CellValues.String;
                cellValue = new CellValue(dataCell.ToString());
            }

            var cell = new Cell
            {
                DataType = excelCellType,
                CellValue = cellValue,
            };

            if (styleIndex != null)
                cell.StyleIndex = styleIndex;

            return cell;
        }

        public static ExcelCreatorWorkbook CreateWorkbook()
        {
            return new ExcelCreatorWorkbook();
        }

        public static ExcelCreatorWorkbook OpenFromTemplate()
        {
            throw new NotImplementedException();
        }

        private enum NumberFormats : uint
        {
            General = 0,
            Zero = 1,
            DecimalTwoDecimals = 2,
            Percentage = 9,
            PercentageTwoDecimals = 10,
            DateTime = 14,
        }
    }
}
