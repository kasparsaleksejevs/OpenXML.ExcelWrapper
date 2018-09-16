using OpenXML.ExcelWrapper.Styling;
using System;
using System.Text.RegularExpressions;

namespace OpenXML.ExcelWrapper
{
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
            if (match is null)
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
}
