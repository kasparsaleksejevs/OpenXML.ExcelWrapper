using System;
using System.Collections.Generic;


namespace OpenXML.ExcelWrapper.old
{
    public class ExcelCreatorSheet
    {
        public string SheetName { get; private set; }

        public SortedList<string, ExcelCell> Cells = new SortedList<string, ExcelCell>();

        public ExcelCreatorSheet(string sheetName)
        {
            this.SheetName = sheetName;
        }

        public void AddCell(string column, int row, int value, CellFormatEnum? cellFormat = null)
        {
            this.AddCellWithValue(column, row, value, cellFormat);
        }

        public void AddCell(string column, int row, int? value, CellFormatEnum? cellFormat = null)
        {
            this.AddCellWithValue(column, row, value, cellFormat);
        }

        public void AddCell(string column, int row, decimal value, CellFormatEnum? cellFormat = null)
        {
            this.AddCellWithValue(column, row, value, cellFormat);
        }

        public void AddCell(string column, int row, decimal? value, CellFormatEnum? cellFormat = null)
        {
            this.AddCellWithValue(column, row, value, cellFormat);
        }

        public void AddCell(string column, int row, string value, CellFormatEnum? cellFormat = null)
        {
            this.AddCellWithValue(column, row, value, cellFormat);
        }

        public void AddCell(string column, int row, DateTime value, CellFormatEnum? cellFormat = null)
        {
            if (cellFormat is null)
                cellFormat = CellFormatEnum.DateTime;
            this.AddCellWithValue(column, row, value, cellFormat);
        }

        public void AddCell(string column, int row, DateTime? value, CellFormatEnum? cellFormat = null)
        {
            if (cellFormat is null)
                cellFormat = CellFormatEnum.DateTime;
            this.AddCellWithValue(column, row, value, cellFormat);
        }

        public void AddCell(string column, int row, float value, CellFormatEnum? cellFormat = null)
        {
            this.AddCellWithValue(column, row, value, cellFormat);
        }

        public void AddCell(string column, int row, float? value, CellFormatEnum? cellFormat = null)
        {
            this.AddCellWithValue(column, row, value, cellFormat);
        }

        public void AddCell(string column, int row, double value, CellFormatEnum? cellFormat = null)
        {
            this.AddCellWithValue(column, row, value, cellFormat);
        }

        public void AddCell(string column, int row, double? value, CellFormatEnum? cellFormat = null)
        {
            this.AddCellWithValue(column, row, value, cellFormat);
        }

        public void AddCell(string column, int row, bool value, CellFormatEnum? cellFormat = null)
        {
            this.AddCellWithValue(column, row, value, cellFormat);
        }

        public void AddCell(string column, int row, bool? value, CellFormatEnum? cellFormat = null)
        {
            this.AddCellWithValue(column, row, value, cellFormat);
        }

        public void AddFormula(string column, int row, string formula, CellFormatEnum? cellFormat = null)
        {
            this.AddCellWithValue(column, row, new ExcelCellFormula(formula), cellFormat);
        }

        private void AddCellWithValue(string column, int row, object value, CellFormatEnum? cellFormat)
        {
            var cell = new ExcelCell(column, row, value, cellFormat);
            this.Cells.Add(cell.Address, cell);
        }
    }
}
