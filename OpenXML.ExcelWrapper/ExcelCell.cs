using OpenXML.ExcelWrapper.Styling;
using System;
using System.Text.RegularExpressions;

namespace OpenXML.ExcelWrapper
{
    /// <summary>
    /// Cell class.
    /// </summary>
    public class ExcelCell
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelCell"/> class.
        /// </summary>
        /// <param name="column">The column letters (A, B, C, etc).</param>
        /// <param name="row">The row number (starting with 1).</param>
        /// <param name="value">The value to insert into the cell.</param>
        /// <param name="cellFormat">The cell data format.</param>
        public ExcelCell(string column, int row, object value, CellFormatEnum? cellFormat = null)
        {
            this.Init(column, row, value, cellFormat);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelCell"/> class.
        /// </summary>
        /// <param name="column">The column number (starting with 1).</param>
        /// <param name="row">The row number (starting with 1).</param>
        /// <param name="value">The value to insert into the cell.</param>
        /// <param name="cellFormat">The cell data format.</param>
        public ExcelCell(int column, int row, object value, CellFormatEnum? cellFormat = null)
        {
            this.Init(GetColumnLetters(column), row, value, cellFormat);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelCell" /> class.
        /// </summary>
        /// <param name="address">The cell address (e.g, 'A1' or 'AZ22').</param>
        /// <param name="value">The value to insert into the cell.</param>
        /// <param name="cellFormat">The cell data format.</param>
        /// <exception cref="ArgumentException">Invalid cell address.</exception>
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

        /// <summary>
        /// Initializes the specified cell.
        /// </summary>
        /// <param name="column">The column letters.</param>
        /// <param name="row">The row number.</param>
        /// <param name="value">The value.</param>
        /// <param name="cellFormat">The cell data format.</param>
        private void Init(string column, int row, object value, CellFormatEnum? cellFormat = null)
        {
            this.Column = column;
            this.Row = row;
            this.Value = this.CheckValueIsFormula(value);
            this.ValueType = value.GetType();
            if (cellFormat.HasValue)
                this.CellStyle = new ExcelCellStyle { CellFormat = cellFormat };
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

        public ExcelCellStyle CellStyle { get; set; }

        /// <summary>
        /// Gets a value indicating whether this cell contains a formula.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance is formula; otherwise, <c>false</c>.
        /// </value>
        internal bool IsFormula { get; private set; }

        internal Type ValueType { get; private set; }

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

        /// <summary>
        /// Checks whether the value is a formula string. If it is - performs additional processing to be compatible with OpenXML.
        /// Formula is specified as "=SUM(A1:B2)". 
        /// Formula string (that should not be calculated) is specified as "'=SUM(A1:B2)".
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>Processed formula (if it is a formula string), otherwise - object is unchanged.</returns>
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
