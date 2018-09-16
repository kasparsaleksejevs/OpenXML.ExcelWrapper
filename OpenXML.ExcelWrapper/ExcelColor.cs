namespace OpenXML.ExcelWrapper
{
    public class ExcelColor
    {
        /// <summary>
        /// Gets or sets the color hexadecimal code (RGB values, e.g., "00FF00" for green color).
        /// </summary>
        /// <value>
        /// The color hexadecimal code.
        /// </value>
        public string ColorHexCode { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelColor"/> class.
        /// </summary>
        /// <param name="colorHexCode">The color hexadecimal code (RGB or ARGB).</param>
        public ExcelColor(string colorHexCode)
        {
            this.ColorHexCode = this.ConvertToArgbColorHex(colorHexCode);
        }

#if NET45 || NET46
        public ExcelColor(System.Drawing.Color color)
        {
            var colorHex = color.A.ToString("X2") + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
            this.ColorHexCode = colorHex;
        }
#endif

        private string ConvertToArgbColorHex(string colorHex)
        {
            var result = colorHex;
            if (colorHex.Length == 6)
                result = "FF" + colorHex;

            return result;
        }

        public override bool Equals(object obj)
        {
            if (obj is null)
                return false;

            if (obj is ExcelColor otherColor)
                return this.ColorHexCode == otherColor.ColorHexCode;

            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            return this.ColorHexCode.GetHashCode();
        }

        public static bool operator ==(ExcelColor obj1, ExcelColor obj2)
        {
            return obj1?.ColorHexCode == obj2?.ColorHexCode;
        }

        public static bool operator !=(ExcelColor obj1, ExcelColor obj2)
        {
            return obj1?.ColorHexCode != obj2?.ColorHexCode;
        }
    }
}
