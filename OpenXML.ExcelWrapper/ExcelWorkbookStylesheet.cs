using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace OpenXML.ExcelWrapper
{
    internal class ExcelWorkbookStylesheet
    {
        private readonly List<Fill> fills = new List<Fill> { new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }, new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } } };
        private readonly List<Font> fonts = new List<Font> { new Font() };
        private readonly List<Border> borders = new List<Border> { new Border() };

        public Stylesheet CreateStylesheet(List<ExcelSheet> sheets)
        {
            var formats = new CellFormats();
            var allStyles = new List<ExcelCellStyle>();
            foreach (var item in sheets.SelectMany(s => s.Cells))
            {
                if (item.Value.CellStyle != null)
                    allStyles.Add(item.Value.CellStyle);
            }

            var distinctStyles = allStyles.Distinct().ToList();
            foreach (var item in distinctStyles)
            {
                var formatIndex = (uint)formats.Count();
                foreach (var style in allStyles.Where(m => m.Equals(item)))
                    style.SetStyleIndex(formatIndex);

                var cellFormat = new CellFormat();
                this.AddNumberFormatToCellFormat(item, cellFormat);
                this.AddBackgroundToCellFormat(item, cellFormat);
                this.AddFontToCellFormat(item, cellFormat);
                this.AddBordersToCellFormat(item, cellFormat);

                formats.AppendChild(cellFormat);
            }

            var styleSheet = new Stylesheet
            {
                Fonts = new Fonts(this.fonts),
                Fills = new Fills(this.fills),
                Borders = new Borders(this.borders),
                CellStyleFormats = new CellStyleFormats(new CellFormat()),
                CellFormats = formats
            };

            return styleSheet;
        }

        private static string ColorToRgbString(System.Drawing.Color color)
        {
            return color.A.ToString("X2") + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
        }

        private void AddNumberFormatToCellFormat(ExcelCellStyle item, CellFormat cellFormat)
        {
            if (item.CellFormat.HasValue)
            {
                cellFormat.NumberFormatId = (uint)item.CellFormat.Value;
                cellFormat.ApplyNumberFormat = true;
            }
        }

        private void AddBackgroundToCellFormat(ExcelCellStyle item, CellFormat cellFormat)
        {
            if (item.BackgroundColor.HasValue)
            {
                var color = ColorToRgbString(item.BackgroundColor.Value);
                var fill = new Fill
                {
                    PatternFill = new PatternFill
                    {
                        ForegroundColor = new ForegroundColor { Rgb = color },
                        PatternType = PatternValues.Solid
                    }
                };

                cellFormat.FillId = GetElementIndex(this.fills, fill);
                cellFormat.ApplyFill = true;
            }
        }

        private void AddFontToCellFormat(ExcelCellStyle item, CellFormat cellFormat)
        {
            if (item.FontColor.HasValue)
            {
                var color = ColorToRgbString(item.FontColor.Value);

                var font = new Font
                {
                    Color = new Color { Rgb = color },
                };

                cellFormat.FontId = GetElementIndex(this.fonts, font);
                cellFormat.ApplyFont = true;
            }
        }

        private void AddBordersToCellFormat(ExcelCellStyle item, CellFormat cellFormat)
        {
            if (item.Borders.HasValue)
            {
                var border = new Border();

                if ((item.Borders.Value & ExcelCellBorderEnum.Top) == ExcelCellBorderEnum.Top)
                    border.AppendChild(new TopBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thick });
                if ((item.Borders.Value & ExcelCellBorderEnum.Bottom) == ExcelCellBorderEnum.Bottom)
                    border.AppendChild(new BottomBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thick });
                if ((item.Borders.Value & ExcelCellBorderEnum.Left) == ExcelCellBorderEnum.Left)
                    border.Append(new LeftBorder());
                if ((item.Borders.Value & ExcelCellBorderEnum.Right) == ExcelCellBorderEnum.Right)
                    border.Append(new RightBorder());
                if ((item.Borders.Value & ExcelCellBorderEnum.Diagonal) == ExcelCellBorderEnum.Diagonal)
                    border.Append(new DiagonalBorder());

                cellFormat.BorderId = GetElementIndex(this.borders, border);
                cellFormat.ApplyBorder = true;
            }
        }

        private uint GetElementIndex<T>(IList<T> collection, T element)
        {
            uint id = 0;
            if (!collection.Contains(element))
            {
                id = (uint)collection.Count();
                collection.Add(element);
            }
            else
                id = (uint)collection.IndexOf(element);

            return id;
        }
    }
}
