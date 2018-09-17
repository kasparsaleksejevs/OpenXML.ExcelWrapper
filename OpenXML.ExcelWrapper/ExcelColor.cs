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
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelColor"/> class.
        /// </summary>
        /// <param name="color">Color info.</param>
        public ExcelColor(System.Drawing.Color color)
        {
            var colorHex = color.A.ToString("X2") + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
            this.ColorHexCode = colorHex;
        }
#endif

        /// <summary>
        /// Converts the hexadecimal RGB color to the hexadecimal ARGB color.
        /// </summary>
        /// <param name="colorHex">The color hexadecimal.</param>
        /// <returns></returns>
        private string ConvertToArgbColorHex(string colorHex)
        {
            var result = colorHex;
            if (colorHex.Length == 6)
                result = "FF" + colorHex;

            return result;
        }

        /// <summary>
        /// Determines whether the specified <see cref="System.Object" />, is equal to this instance.
        /// </summary>
        /// <param name="obj">The <see cref="System.Object" /> to compare with this instance.</param>
        /// <returns>
        ///   <c>true</c> if the specified <see cref="System.Object" /> is equal to this instance; otherwise, <c>false</c>.
        /// </returns>
        public override bool Equals(object obj)
        {
            if (obj is null)
                return false;

            if (obj is ExcelColor otherColor)
                return this.ColorHexCode == otherColor.ColorHexCode;

            return base.Equals(obj);
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            return this.ColorHexCode.GetHashCode();
        }

        /// <summary>
        /// Implements the operator ==.
        /// </summary>
        /// <param name="obj1">The obj1.</param>
        /// <param name="obj2">The obj2.</param>
        /// <returns>
        /// The result of the operator.
        /// </returns>
        public static bool operator ==(ExcelColor obj1, ExcelColor obj2)
        {
            return obj1?.ColorHexCode == obj2?.ColorHexCode;
        }

        /// <summary>
        /// Implements the operator !=.
        /// </summary>
        /// <param name="obj1">The obj1.</param>
        /// <param name="obj2">The obj2.</param>
        /// <returns>
        /// The result of the operator.
        /// </returns>
        public static bool operator !=(ExcelColor obj1, ExcelColor obj2)
        {
            return obj1?.ColorHexCode != obj2?.ColorHexCode;
        }
    }
}
