namespace OpenXML.ExcelWrapper
{
    public class ExcelCellStyle
    {
        public CellFormatEnum? CellFormat { get; set; }

        public System.Drawing.Color? FontColor { get; set; }

        public System.Drawing.Color? BackgroundColor { get; set; }

        public ExcelCellBorderEnum? Borders { get; set; }

        internal uint StyleIndex { get; private set; }

        // borders
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
            hash ^= 8 * this.Borders.GetHashCode();
            return hash;
        }

        internal void SetStyleIndex(uint styleIndex)
        {
            this.StyleIndex = styleIndex;
        }
    }
}
