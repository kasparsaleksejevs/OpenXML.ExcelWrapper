using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenXML.ExcelWrapper.Styling
{
    public class ExcelCellStyle : IEquatable<ExcelCellStyle>
    {
        public CellFormatEnum? CellFormat { get; set; }

        public ExcelColor BackgroundColor { get; set; }

        public ICollection<ExcelCellStyleBorder> Borders { get; set; } = new List<ExcelCellStyleBorder>();

        public ExcelCellStyleFont Font { get; set; }

        public HorizontalAlignmentEnum? HorizontalAlignment { get; set; }

        public VerticalAlignmentEnum? VerticalAlignment { get; set; }

        public bool? WrapText { get; set; }

        public int? TextRotation { get; set; }

        public bool? ShrinkToFit { get; set; }

        internal uint StyleIndex { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelCellStyle"/> class.
        /// </summary>
        public ExcelCellStyle()
        {
        }

        /// <summary>
        /// Indicates whether the current object is equal to another object of the same type.
        /// </summary>
        /// <param name="other">An object to compare with this object.</param>
        /// <returns>
        /// true if the current object is equal to the <paramref name="other" /> parameter; otherwise, false.
        /// </returns>
        /// <exception cref="NotImplementedException"></exception>
        public bool Equals(ExcelCellStyle other)
        {
            if (other is null)
                return false;

            if (ReferenceEquals(this, other))
                return true;

            var isEqual = this.CellFormat == other.CellFormat
                    && this.BackgroundColor == other.BackgroundColor
                    && this.Borders.SequenceEqual(other.Borders)
                    && (this.Font is null && other.Font is null || this.Font.Equals(other.Font))
                    && this.HorizontalAlignment == other.HorizontalAlignment
                    && this.VerticalAlignment == other.VerticalAlignment
                    && this.WrapText == other.WrapText
                    && this.TextRotation == other.TextRotation
                    && this.ShrinkToFit == other.ShrinkToFit;

            return isEqual;
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

            if (obj is ExcelCellStyle otherStyle)
                return this.Equals(otherStyle);

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
            // ToDo: this hashing algorithm kind of sucks, maybe there is something better?
            int hash = 0;
            hash ^= 1 * this.CellFormat?.GetHashCode() ?? 1;
            hash ^= 2 * this.BackgroundColor?.GetHashCode() ?? 1;
            hash ^= 4 * this.Borders.Select(s => s.GetHashCode()).Sum().GetHashCode();
            hash ^= 8 * this.Font?.GetHashCode() ?? 1;
            hash ^= 16 * this.HorizontalAlignment?.GetHashCode() ?? 1;
            hash ^= 32 * this.VerticalAlignment?.GetHashCode() ?? 1;
            hash ^= 64 * this.WrapText?.GetHashCode() ?? 1;
            hash ^= 128 * this.TextRotation?.GetHashCode() ?? 1;
            hash ^= 256 * this.ShrinkToFit?.GetHashCode() ?? 1;

            return hash;
        }

        /// <summary>
        /// Sets the index of the style for this cell.
        /// </summary>
        /// <param name="styleIndex">Index of the style.</param>
        internal void SetStyleIndex(uint styleIndex)
        {
            this.StyleIndex = styleIndex;
        }
    }
}
