using OpenXML.ExcelWrapper.Styling;
using System.Collections.Generic;

namespace OpenXML.ExcelWrapper
{
    /// <summary>
    /// Excel sheet class.
    /// </summary>
    public class ExcelSheet
    {
        /// <summary>
        /// Gets the name of the sheet.
        /// </summary>
        /// <value>
        /// The name of the sheet.
        /// </value>
        public string SheetName { get; private set; }

        /// <summary>
        /// The cells collection to be written to the Excel document.
        /// </summary>
        internal readonly SortedList<string, ExcelCell> Cells = new SortedList<string, ExcelCell>();

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelSheet"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        public ExcelSheet(string name)
        {
            this.SheetName = name;
        }

        /// <summary>
        /// Adds the or updates the cell.
        /// </summary>
        /// <param name="address">The cell address.</param>
        /// <param name="value">The cell data value.</param>
        /// <param name="cellFormat">The cell data format.</param>
        /// <param name="style">The cell style.</param>
        public void AddOrUpdateCell(string address, object value, CellFormatEnum? cellFormat = null, ExcelCellStyle style = null)
        {
            var cell = new ExcelCell(address, value, cellFormat);
            if (style != null)
                cell.CellStyle = style;

            if (this.Cells.ContainsKey(address))
                this.Cells[address] = cell;
            else
                this.Cells.Add(address, cell);
        }

        /// <summary>
        /// Adds the or updates the cell.
        /// </summary>
        /// <param name="cell">The cell.</param>
        public void AddOrUpdateCell(ExcelCell cell)
        {
            var address = cell.Address;

            if (this.Cells.ContainsKey(address))
                this.Cells[address] = cell;
            else
                this.Cells.Add(address, cell);
        }
    }
}
