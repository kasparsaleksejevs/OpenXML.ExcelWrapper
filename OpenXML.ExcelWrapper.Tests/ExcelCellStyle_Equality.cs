using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXML.ExcelWrapper.Styling;
using Shouldly;
using System.Collections.Generic;
using System.Linq;

namespace OpenXML.ExcelWrapper.Tests
{
    [TestClass]
    public class ExcelCellStyle_Equality
    {
        [TestMethod]
        public void ExcelCellStyle_OneInstance()
        {
            var cellStyleA = new ExcelCellStyle { CellFormat = CellFormatEnum.DecimalTwoDecimals, ShrinkToFit = true, BackgroundColor = new ExcelColor("0FA3D1") };
            var cellStyleB = cellStyleA;
            cellStyleA.Equals(cellStyleB).ShouldBeTrue();
        }

        [TestMethod]
        public void ExcelCellStyle_TwoInstances_Identical()
        {
            var cellStyleA = new ExcelCellStyle { CellFormat = CellFormatEnum.DecimalTwoDecimals, ShrinkToFit = true, BackgroundColor = new ExcelColor("0FA3D1") };
            var cellStyleB = new ExcelCellStyle { CellFormat = CellFormatEnum.DecimalTwoDecimals, ShrinkToFit = true, BackgroundColor = new ExcelColor("0FA3D1") };

            cellStyleA.Equals(cellStyleB).ShouldBeTrue();
        }

        [TestMethod]
        public void ExcelCellStyle_ListOfTwoInstancesDistinct_ShouldReturnOneInstance()
        {
            var lst = new List<ExcelCellStyle> {
                new ExcelCellStyle { CellFormat = CellFormatEnum.DecimalTwoDecimals, ShrinkToFit = true, BackgroundColor = new ExcelColor("0FA3D1") },
                new ExcelCellStyle { CellFormat = CellFormatEnum.DecimalTwoDecimals, ShrinkToFit = true, BackgroundColor = new ExcelColor("0FA3D1") }
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
                    Font = new ExcelCellStyleFont { Color = new ExcelColor("123456"), IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                },
                new ExcelCellStyle {
                    CellFormat = CellFormatEnum.DecimalTwoDecimals,
                    Font = new ExcelCellStyleFont { Color = new ExcelColor("123456"), IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
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
                    Font = new ExcelCellStyleFont { Color = new ExcelColor("123456"), IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                    Borders = new List<ExcelCellStyleBorder>{ new ExcelCellStyleBorder( ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Hair, new ExcelColor("559999") ), new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.DashDot, new ExcelColor("667788")) },
                    ShrinkToFit = true,
                    BackgroundColor = new ExcelColor("0FA3D1")
                },
                new ExcelCellStyle {
                    CellFormat = CellFormatEnum.PercentageTwoDecimals,
                    HorizontalAlignment = HorizontalAlignmentEnum.CenterContinuous,
                    TextRotation = 99,
                    VerticalAlignment = VerticalAlignmentEnum.Justify,
                    WrapText = true,
                    Font = new ExcelCellStyleFont { Color = new ExcelColor("123456"), IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                    Borders = new List<ExcelCellStyleBorder>{ new ExcelCellStyleBorder( ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Hair, new ExcelColor("559999") ), new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.DashDot, new ExcelColor("667788")) },
                    ShrinkToFit = true,
                    BackgroundColor = new ExcelColor("0FA3D1")
                },
                new ExcelCellStyle {
                    CellFormat = CellFormatEnum.PercentageTwoDecimals,
                    HorizontalAlignment = HorizontalAlignmentEnum.CenterContinuous,
                    TextRotation = 99,
                    VerticalAlignment = VerticalAlignmentEnum.Justify,
                    WrapText = true,
                    Font = new ExcelCellStyleFont { Color = new ExcelColor("123456"), IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                    Borders = new List<ExcelCellStyleBorder>{ new ExcelCellStyleBorder( ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Hair, new ExcelColor("559999") ), new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.DashDot, new ExcelColor("667788")) },
                    ShrinkToFit = true,
                    BackgroundColor = new ExcelColor("0FA3D1")
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
                    Font = new ExcelCellStyleFont { Color = new ExcelColor("123456"), IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                    Borders = new List<ExcelCellStyleBorder>{ new ExcelCellStyleBorder( ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Hair, new ExcelColor("559999") ), new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.DashDot, new ExcelColor("667788")) },
                    ShrinkToFit = true,
                    BackgroundColor = new ExcelColor("0FA3D1")
                },
                new ExcelCellStyle {
                    CellFormat = CellFormatEnum.PercentageTwoDecimals,
                    HorizontalAlignment = HorizontalAlignmentEnum.CenterContinuous,
                    TextRotation = 99,
                    VerticalAlignment = VerticalAlignmentEnum.Center,
                    WrapText = true,
                    Font = new ExcelCellStyleFont { Color = new ExcelColor("123456"), IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                    Borders = new List<ExcelCellStyleBorder>{ new ExcelCellStyleBorder( ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Hair, new ExcelColor("559999") ), new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.DashDot, new ExcelColor("667788")) },
                    ShrinkToFit = true,
                    BackgroundColor = new ExcelColor("0FA3D1")
                },
                new ExcelCellStyle {
                    CellFormat = CellFormatEnum.PercentageTwoDecimals,
                    HorizontalAlignment = HorizontalAlignmentEnum.CenterContinuous,
                    TextRotation = 99,
                    VerticalAlignment = VerticalAlignmentEnum.Justify,
                    WrapText = true,
                    Font = new ExcelCellStyleFont { Color = new ExcelColor("123456"), IsBold = true, FontName =" test", IsItalic = true, IsUnderline = true, Size = 99 },
                    Borders = new List<ExcelCellStyleBorder>{ new ExcelCellStyleBorder( ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Hair, new ExcelColor("559999") ), new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.DashDot, new ExcelColor("667788")) },
                    ShrinkToFit = true,
                    BackgroundColor = new ExcelColor("0FA3D1")
                },
            };

            lst.Distinct().Count().ShouldBe(2);
        }
    }
}
