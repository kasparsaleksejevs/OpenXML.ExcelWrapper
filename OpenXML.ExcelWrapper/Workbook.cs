using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

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

        private Stylesheet CreateStylesheet()
        {
            var formats = new CellFormats();
            var fills = new List<Fill>();

            var allStyles = new List<ExcelCellStyle>();
            foreach (var item in this.sheets.SelectMany(s => s.Cells))
            {
                if (item.Value.CellStyle != null)
                    allStyles.Add(item.Value.CellStyle);
            }

            var distinctStyles = allStyles.Distinct().ToList();
            foreach (var item in distinctStyles)
            {
                var formatIndex = (uint)formats.Count();
                foreach (var style in allStyles.Where(m => m.Equals(item)))
                    style.SetStyleIndex(formatIndex);

                var cellFormat = new CellFormat();
                if (item.CellFormat.HasValue)
                {
                    cellFormat.NumberFormatId = (uint)item.CellFormat.Value;
                    cellFormat.ApplyNumberFormat = true;
                }

                if (item.BackgroundColor.HasValue)
                {
                    var fill = new Fill
                    {
                        PatternFill = new PatternFill
                        {
                            BackgroundColor = new BackgroundColor { Rgb = ColorToRgbString(item.BackgroundColor.Value) },
                            PatternType = PatternValues.Solid
                        }
                    };

                    if (!fills.Contains(fill))
                    {
                        cellFormat.FillId = (uint)fills.Count();
                        fills.Add(fill);
                    }
                    else
                        cellFormat.FillId = (uint)fills.IndexOf(fill);

                    cellFormat.ApplyFill = true;
                    cellFormat.FormatId = 0;
                }

                formats.AppendChild(cellFormat);
            }

            var styleSheet = new Stylesheet
            {
                Fonts = new Fonts(new Font()),
                Fills = new Fills(fills),
                Borders = new Borders(new Border()),
                CellStyleFormats = new CellStyleFormats(new CellFormat()),
                CellFormats = formats
            };

            return styleSheet;
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
                stylesPart.Stylesheet = this.CreateStylesheet();

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
            {
                cell.StyleIndex = excelCell.CellStyle.StyleIndex;
            }


            return cell;
        }

        private static string ColorToRgbString(System.Drawing.Color color)
        {
            return color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
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

    public class ExcelSheet
    {
        public string SheetName { get; private set; }

        public readonly SortedList<string, ExcelCell> Cells = new SortedList<string, ExcelCell>();

        public ExcelSheet(string name)
        {
            this.SheetName = name;
        }

        public void AddOrUpdateCell(ExcelCell cell)
        {
            var address = cell.Address;

            if (this.Cells.ContainsKey(address))
                this.Cells[address] = cell;
            else
                this.Cells.Add(address, cell);
        }
    }

    public class ExcelCell
    {
        public ExcelCell(string column, int row, object value, CellFormatEnum? cellFormat = null)
        {
            this.Init(column, row, value, cellFormat);
        }

        public ExcelCell(int column, int row, object value, CellFormatEnum? cellFormat = null)
        {
            this.Init(GetColumnLetters(column), row, value, cellFormat);
        }

        public ExcelCell(string address, object value, CellFormatEnum? cellFormat = null)
        {
            var addressRegex = new Regex(@"(?<col>([A-Z]|[a-z])+)(?<row>([1-9]\d*)+)");
            var match = addressRegex.Match(address);
            if (match == null)
                throw new ArgumentException($"Invalid cell address {address}");

            var column = match.Groups["col"].Value;
            var row = Convert.ToInt32(match.Groups["row"].Value);
            this.Init(column, row, value, cellFormat);
        }

        private void Init(string column, int row, object value, CellFormatEnum? cellFormat = null)
        {
            this.Column = column;
            this.Row = row;
            this.Value = this.CheckValueIsFormula(value);
            this.ValueType = value.GetType();
            if (cellFormat.HasValue)
                this.CellStyle = new ExcelCellStyle { CellFormat = cellFormat };
        }

        public bool IsFormula { get; private set; }

        public int Row { get; private set; }

        public string Column { get; private set; }

        public object Value { get; private set; }

        public Type ValueType { get; private set; }

        public string Address
        {
            get
            {
                return $"{this.Column}{this.Row}";
            }
        }

        public ExcelCellStyle CellStyle { get; set; }

        /// <summary>
        /// Gets the column letters from numeric index.
        /// E.g., 1=A, 1466 = BDJ.
        /// </summary>
        /// <param name="columnNr">The column nr (1-based).</param>
        public static string GetColumnLetters(int columnNr)
        {
            // 1 = A, 256 = IV, 419  = PC, 1466 = BDJ
            const int letterCount = 26;
            const int letterCount2 = letterCount * letterCount;
            const int baseLetter = 'A' - 1;

            var letter3 = columnNr / letterCount2;
            var letter3Rem = columnNr % letterCount2;

            var letter2 = letter3Rem / letterCount;
            var letter1 = letter3Rem % letterCount;

            var result = "";
            if (letter3 > 0)
                result += (char)(baseLetter + letter3);

            if (letter2 > 0)
                result += (char)(baseLetter + letter2);

            result += (char)(baseLetter + letter1);

            return result;
        }

        private object CheckValueIsFormula(object value)
        {
            if (value is string stringValue)
            {
                // it is a formula
                if (stringValue.StartsWith("="))
                {
                    this.IsFormula = true;
                    stringValue = stringValue.Substring(1);
                }

                // it is a string that looks like a formula
                if (stringValue.StartsWith("'="))
                {
                    this.IsFormula = false;
                    stringValue = stringValue.Substring(1);
                }

                return stringValue;
            }

            return value;
        }
    }

    public enum CellFormatEnum : uint
    {
        General = 0,
        Zero = 1,
        DecimalTwoDecimals = 2,
        Percentage = 9,
        PercentageTwoDecimals = 10,
        DateTime = 14,
    }

    public class ExcelCellStyle
    {
        public uint StyleIndex { get; private set; }

        public CellFormatEnum? CellFormat { get; set; }

        public System.Drawing.Color? FontColor { get; set; }

        public System.Drawing.Color? BackgroundColor { get; set; }


        // borders
        // color
        // background
        // etc

        public ExcelCellStyle()
        {
        }

        public override bool Equals(object obj)
        {
            if (obj is ExcelCellStyle otherStyle)
            {
                return this.CellFormat == otherStyle.CellFormat
                    && this.BackgroundColor == otherStyle.BackgroundColor
                    && this.FontColor == otherStyle.FontColor;
            }

            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            // ToDo: this hashing algorithm kind of sucks, maybe there is something better?
            int hash = 0;
            hash ^= 1 * this.CellFormat.GetHashCode();
            hash ^= 2 * this.BackgroundColor.GetHashCode();
            hash ^= 4 * this.FontColor.GetHashCode();
            return hash;
        }

        internal void SetStyleIndex(uint styleIndex)
        {
            this.StyleIndex = styleIndex;
        }
    }
}
