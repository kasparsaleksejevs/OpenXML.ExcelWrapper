namespace OpenXML.ExcelWrapper.Styling
{
    public class ExcelCellStyleBorder
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelCellStyleBorder"/> class.
        /// </summary>
        /// <param name="border">The border (top, bottom, etc).</param>
        /// <param name="style">The style of the border.</param>
        /// <param name="color">The color of the border.</param>
        public ExcelCellStyleBorder(ExcelCellBorderEnum border, ExcelCellStyleBorderSizeEnum style = ExcelCellStyleBorderSizeEnum.Thin, System.Drawing.Color? color = null)
        {
            this.Border = border;
            this.Style = style;
            this.Color = color;
        }

        public ExcelCellBorderEnum Border { get; set; }

        public System.Drawing.Color? Color { get; set; }

        public ExcelCellStyleBorderSizeEnum? Style { get; set; }

        /// <summary>
        /// Determines whether the specified <see cref="System.Object" />, is equal to this instance.
        /// </summary>
        /// <param name="obj">The <see cref="System.Object" /> to compare with this instance.</param>
        /// <returns>
        ///   <c>true</c> if the specified <see cref="System.Object" /> is equal to this instance; otherwise, <c>false</c>.
        /// </returns>
        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (obj is ExcelCellStyleBorder otherBorder)
            {
                return this.Border == otherBorder.Border
                    && this.Style == otherBorder.Style
                    && this.Color == otherBorder.Color;
            }

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
            int hash = 0;
            hash ^= 1 * this.Border.GetHashCode();
            hash ^= 8 * this.Style?.GetHashCode() ?? 1;
            hash ^= 16 * this.Color?.GetHashCode() ?? 1;

            return hash;
        }
    }
}
