using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXML.ExcelWrapper.Styling;
using Shouldly;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace OpenXML.ExcelWrapper.Tests
{
    [TestClass]
    public class ExcelCellStyle_Equality
    {
        [TestMethod]
        public void ExcelCellStyle_OneInstance()
        {
            var cellStyleA = new ExcelCellStyle { CellFormat = CellFormatEnum.DecimalTwoDecimals, ShrinkToFit = true, BackgroundColor = Color.FloralWhite };
            var cellStyleB = cellStyleA;
            cellStyleA.Equals(cellStyleB).ShouldBeTrue();
        }

        [TestMethod]
        public void ExcelCellStyle_TwoInstances_Identical()
        {
            var cellStyleA = new ExcelCellStyle { CellFormat = CellFormatEnum.DecimalTwoDecimals, ShrinkToFit = true, BackgroundColor = Color.FloralWhite };
            var cellStyleB = new ExcelCellStyle { CellFormat = CellFormatEnum.DecimalTwoDecimals, ShrinkToFit = true, BackgroundColor = Color.FloralWhite };

            cellStyleA.Equals(cellStyleB).ShouldBeTrue();
        }

        [TestMethod]
        public void ExcelCellStyle_ListOfTwoInstancesDistinct_ShouldReturnOneInstance()
        {
            var lst = new List<ExcelCellStyle> {
                new ExcelCellStyle { CellFormat = CellFormatEnum.DecimalTwoDecimals, ShrinkToFit = true, BackgroundColor = Color.FloralWhite },
                new ExcelCellStyle { CellFormat = CellFormatEnum.DecimalTwoDecimals, ShrinkToFit = true, BackgroundColor = Color.FloralWhite }
            };

            lst[0].ShouldBe(lst[1]);
            lst[0].GetHashCode().ShouldBe(lst[1].GetHashCode());
            lst.Distinct().Count().ShouldBe(1);
        }

        [TestMethod]
        public void ExcelCellStyle_ListOfTwoInstancesWithFontsDistinct_ShouldReturnOneInstance()
        {
            var lst = new List<ExcelCellStyle> {
                new ExcelCellStyle {
                    CellFormat = CellFormatEnum.DecimalTwoDecimals,
                    Font = new ExcelCellStyleFont { Color = Color.Fuchsia, IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                },
                new ExcelCellStyle {
                    CellFormat = CellFormatEnum.DecimalTwoDecimals,
                    Font = new ExcelCellStyleFont { Color = Color.Fuchsia, IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                },
            };

            lst[0].GetHashCode().ShouldBe(lst[1].GetHashCode());
            lst[0].ShouldBe(lst[1]);
            lst.Distinct().Count().ShouldBe(1);
        }

        [TestMethod]
        public void ExcelCellStyle_ListOfThreeInstancesDistinct_ShouldReturnOneInstance()
        {
            var lst = new List<ExcelCellStyle> {
                new ExcelCellStyle {
                    CellFormat = CellFormatEnum.PercentageTwoDecimals,
                    HorizontalAlignment = HorizontalAlignmentEnum.CenterContinuous,
                    TextRotation = 99,
                    VerticalAlignment = VerticalAlignmentEnum.Justify,
                    WrapText = true,
                    Font = new ExcelCellStyleFont { Color = Color.Fuchsia, IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                    Borders = new List<ExcelCellStyleBorder>{ new ExcelCellStyleBorder( ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Hair, Color.Aqua ), new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.DashDot, Color.Bisque) },
                    ShrinkToFit = true,
                    BackgroundColor = Color.FloralWhite
                },
                new ExcelCellStyle {
                    CellFormat = CellFormatEnum.PercentageTwoDecimals,
                    HorizontalAlignment = HorizontalAlignmentEnum.CenterContinuous,
                    TextRotation = 99,
                    VerticalAlignment = VerticalAlignmentEnum.Justify,
                    WrapText = true,
                    Font = new ExcelCellStyleFont { Color = Color.Fuchsia, IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                    Borders = new List<ExcelCellStyleBorder>{ new ExcelCellStyleBorder( ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Hair, Color.Aqua ), new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.DashDot, Color.Bisque) },
                    ShrinkToFit = true,
                    BackgroundColor = Color.FloralWhite
                },
                new ExcelCellStyle {
                    CellFormat = CellFormatEnum.PercentageTwoDecimals,
                    HorizontalAlignment = HorizontalAlignmentEnum.CenterContinuous,
                    TextRotation = 99,
                    VerticalAlignment = VerticalAlignmentEnum.Justify,
                    WrapText = true,
                    Font = new ExcelCellStyleFont { Color = Color.Fuchsia, IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                    Borders = new List<ExcelCellStyleBorder>{ new ExcelCellStyleBorder( ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Hair, Color.Aqua ), new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.DashDot, Color.Bisque) },
                    ShrinkToFit = true,
                    BackgroundColor = Color.FloralWhite
                },
            };

            lst.Distinct().Count().ShouldBe(1);
        }

        [TestMethod]
        public void ExcelCellStyle_ListOfThreeInstancesDistinct_ShouldReturnTwoInstances()
        {
            var lst = new List<ExcelCellStyle> {
                new ExcelCellStyle {
                    CellFormat = CellFormatEnum.PercentageTwoDecimals,
                    HorizontalAlignment = HorizontalAlignmentEnum.CenterContinuous,
                    TextRotation = 99,
                    VerticalAlignment = VerticalAlignmentEnum.Justify,
                    WrapText = true,
                    Font = new ExcelCellStyleFont { Color = Color.Fuchsia, IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                    Borders = new List<ExcelCellStyleBorder>{ new ExcelCellStyleBorder( ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Hair, Color.Aqua ), new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.DashDot, Color.Bisque) },
                    ShrinkToFit = true,
                    BackgroundColor = Color.FloralWhite
                },
                new ExcelCellStyle {
                    CellFormat = CellFormatEnum.PercentageTwoDecimals,
                    HorizontalAlignment = HorizontalAlignmentEnum.CenterContinuous,
                    TextRotation = 99,
                    VerticalAlignment = VerticalAlignmentEnum.Center,
                    WrapText = true,
                    Font = new ExcelCellStyleFont { Color = Color.Fuchsia, IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                    Borders = new List<ExcelCellStyleBorder>{ new ExcelCellStyleBorder( ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Hair, Color.Aqua ), new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.DashDot, Color.Bisque) },
                    ShrinkToFit = true,
                    BackgroundColor = Color.FloralWhite
                },
                new ExcelCellStyle {
                    CellFormat = CellFormatEnum.PercentageTwoDecimals,
                    HorizontalAlignment = HorizontalAlignmentEnum.CenterContinuous,
                    TextRotation = 99,
                    VerticalAlignment = VerticalAlignmentEnum.Justify,
                    WrapText = true,
                    Font = new ExcelCellStyleFont { Color = Color.Fuchsia, IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                    Borders = new List<ExcelCellStyleBorder>{ new ExcelCellStyleBorder( ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Hair, Color.Aqua ), new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.DashDot, Color.Bisque) },
                    ShrinkToFit = true,
                    BackgroundColor = Color.FloralWhite
                },
            };

            lst.Distinct().Count().ShouldBe(2);
        }
    }
}
