namespace OpenXML.ExcelWrapper.Styling
{
    public class ExcelCellStyleFont
    {
        public string FontName { get; set; }

        public double? Size { get; set; }

        public System.Drawing.Color? Color { get; set; }

        public bool IsBold { get; set; }

        public bool IsItalic { get; set; }

        public bool IsUnderline { get; set; }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (obj is ExcelCellStyleFont FontName)
            {
                return this.FontName == FontName.FontName
                    && this.Size == FontName.Size
                    && this.Color == FontName.Color
                    && this.IsBold == FontName.IsBold
                    && this.IsItalic == FontName.IsItalic
                    && this.IsUnderline == FontName.IsUnderline;
            }

            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            int hash = 0;
            hash ^= 1 * this.FontName?.GetHashCode() ?? 1;
            hash ^= 8 * this.Size?.GetHashCode() ?? 1;
            hash ^= 16 * this.Color?.ToArgb().GetHashCode() ?? 1;
            hash ^= 32 * this.IsBold.GetHashCode();
            hash ^= 64 * this.IsItalic.GetHashCode();
            hash ^= 128 * this.IsUnderline.GetHashCode();

            return hash;
        }
    }
}
