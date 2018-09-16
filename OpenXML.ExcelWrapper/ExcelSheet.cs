using System.Collections.Generic;

namespace OpenXML.ExcelWrapper
{
    public class ExcelSheet
    {
        public string SheetName { get; private set; }

        public readonly SortedList<string, ExcelCell> Cells = new SortedList<string, ExcelCell>();

        public ExcelSheet(string name)
        {
            this.SheetName = name;
        }

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
