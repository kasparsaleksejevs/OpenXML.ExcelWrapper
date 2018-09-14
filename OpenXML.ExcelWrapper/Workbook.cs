using System;
using System.Collections.Generic;
using System.IO;
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

        }

        public byte[] Save()
        {

            return null;
        }

        public static ExcelWorkbook FromFile(string fileName)
        {
            throw new NotImplementedException();
        }

        public static ExcelWorkbook FromStream(Stream stream)
        {
            throw new NotImplementedException();
        }
    }

    public class ExcelSheet
    {
        public string SheetName { get; private set; }

        private readonly SortedList<string, ExcelCell> cells = new SortedList<string, ExcelCell>();

        public ExcelSheet(string name)
        {
            this.SheetName = name;
        }

        public void AddOrUpdateCell(ExcelCell cell)
        {
            var address = cell.Address;

            if (this.cells.ContainsKey(address))
                this.cells[address] = cell;
            else
                this.cells.Add(address, cell);
        }
    }

    public class ExcelCell
    {
        public ExcelCell(string address, object value)
        {
            var addressRegex = new Regex(@"(?<col>([A-Z]|[a-z])+)(?<row>([1-9]\d*)+)");
            var match = addressRegex.Match(address);
            if (match == null)
                throw new ArgumentException($"Invalid cell address {address}");

            this.Column = match.Groups["col"].Value;
            this.Row = Convert.ToInt32(match.Groups["row"].Value);
            this.Value = value;
        }

        public ExcelCell(int column, int row, object value)
        {
            this.Column = GetColumnLetters(column);
            this.Row = row;
            this.Value = value;
        }

        public ExcelCell(string column, int row, object value)
        {
            this.Column = column;
            this.Row = row;
            this.Value = value;
        }

        public int Row { get; private set; }

        public string Column { get; private set; }

        public object Value { get; private set; }

        public string Address
        {
            get
            {
                return $"{this.Column}{this.Row}";
            }
        }

        public CellFormatEnum CellFormat { get; set; }

        public CellStyle CellStyle { get; set; }

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
    }

    public enum CellFormatEnum : uint
    {
        General = 0,
        DateTime = 1,
        DecimalTwoDecimals = 2,
        Percentage = 3,
        PercentageTwoDecimals = 4
    }

    public class CellStyle
    {
        // borders
        // color
        // background
        // etc
    }
}
