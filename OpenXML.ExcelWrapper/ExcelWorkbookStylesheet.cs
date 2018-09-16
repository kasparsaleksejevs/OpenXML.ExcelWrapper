using DocumentFormat.OpenXml.Spreadsheet;
using OpenXML.ExcelWrapper.Styling;
using System;
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
                this.AddAlignmentToCellFormat(item, cellFormat);
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

                cellFormat.FillId = GetElementIndexOrAdd(this.fills, fill);
                cellFormat.ApplyFill = true;
            }
        }

        private void AddFontToCellFormat(ExcelCellStyle item, CellFormat cellFormat)
        {
            if (item.Font != null)
            {
                var color = ColorToRgbString(item.Font.Color.Value);

                var font = new Font();
                if (!string.IsNullOrWhiteSpace(item.Font.FontName))
                    font.AppendChild(new FontName { Val = item.Font.FontName });

                if (item.Font.Size.HasValue)
                    font.AppendChild(new FontSize { Val = item.Font.Size.Value });

                if (item.Font.Color.HasValue)
                    font.AppendChild(new Color { Rgb = color });

                if (item.Font.IsBold)
                    font.AppendChild(new Bold { Val = true });

                if (item.Font.IsItalic)
                    font.AppendChild(new Italic { Val = true });

                if (item.Font.IsUnderline)
                    font.AppendChild(new Underline { Val = UnderlineValues.Single });

                cellFormat.FontId = GetElementIndexOrAdd(this.fonts, font);
                cellFormat.ApplyFont = true;
            }
        }

        private void AddAlignmentToCellFormat(ExcelCellStyle item, CellFormat cellFormat)
        {
            var alignment = new Alignment() { WrapText = false, TextRotation = 99, ShrinkToFit = true, };
            if (item.HorizontalAlignment.HasValue)
                alignment.Horizontal = (HorizontalAlignmentValues)(int)item.HorizontalAlignment.Value;

            if (item.VerticalAlignment.HasValue)
                alignment.Vertical = (VerticalAlignmentValues)(int)item.VerticalAlignment.Value;

            if (item.WrapText.HasValue)
                alignment.WrapText = item.WrapText.Value;

            if (item.TextRotation.HasValue)
                alignment.TextRotation = (uint)item.TextRotation.Value;

            if (item.ShrinkToFit.HasValue)
                alignment.ShrinkToFit = item.ShrinkToFit.Value;

            cellFormat.ApplyAlignment = true;
        }

        private void AddBordersToCellFormat(ExcelCellStyle item, CellFormat cellFormat)
        {
            if (item.Borders != null && item.Borders.Count > 0)
            {
                var border = new Border();
                foreach (var borderItem in item.Borders)
                {
                    var color = new Color();
                    if (borderItem.Color is null)
                        color.Auto = true;
                    else
                        color.Rgb = ColorToRgbString(borderItem.Color.Value);

                    switch (borderItem.Border)
                    {
                        case ExcelCellBorderEnum.None:
                            break;
                        case ExcelCellBorderEnum.Left:
                            border.AppendChild(new LeftBorder { Color = color, Style = (BorderStyleValues)(int)borderItem.Style });
                            break;
                        case ExcelCellBorderEnum.Right:
                            border.AppendChild(new RightBorder { Color = color, Style = (BorderStyleValues)(int)borderItem.Style });
                            break;
                        case ExcelCellBorderEnum.Top:
                            border.AppendChild(new TopBorder { Color = color, Style = (BorderStyleValues)(int)borderItem.Style });
                            break;
                        case ExcelCellBorderEnum.Bottom:
                            border.AppendChild(new BottomBorder { Color = color, Style = (BorderStyleValues)(int)borderItem.Style });
                            break;
                        case ExcelCellBorderEnum.Diagonal:
                            border.AppendChild(new DiagonalBorder { Color = color, Style = (BorderStyleValues)(int)borderItem.Style });
                            break;
                        default:
                            throw new NotSupportedException();
                    }
                }

                cellFormat.BorderId = GetElementIndexOrAdd(this.borders, border);
                cellFormat.ApplyBorder = true;
            }
        }

        private uint GetElementIndexOrAdd<T>(IList<T> collection, T element)
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
