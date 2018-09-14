namespace OpenXML.ExcelWrapper.old
{
    public class ExcelCell
    {
        public ExcelCell(string column, int row, object value, CellFormatEnum? cellFormat = null)
        {
            this.Column = column;
            this.Row = row;
            this.Value = value;

            if (cellFormat.HasValue)
                this.CellFormatId = (uint)cellFormat.Value;
        }

        public int Row { get; set; }

        public string Column { get; set; }

        public object Value { get; set; }

        public string Address
        {
            get
            {
                return $"{this.Column}{this.Row}";
            }
        }

        public uint? CellFormatId { get; set; }

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
}
