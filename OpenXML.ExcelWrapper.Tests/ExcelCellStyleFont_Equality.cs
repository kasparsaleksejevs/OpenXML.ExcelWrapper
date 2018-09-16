using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXML.ExcelWrapper.Styling;
using Shouldly;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace OpenXML.ExcelWrapper.Tests
{
    [TestClass]
    public class ExcelCellStyleFont_Equality
    {
        [TestMethod]
        public void ExcelCellStyleFont_OneInstance()
        {
            var cellStyleA = new ExcelCellStyleFont { Color = Color.Red, FontName = "test", IsBold = true, IsItalic = false, Size = 99 };
            var cellStyleB = cellStyleA;
            cellStyleA.Equals(cellStyleB).ShouldBeTrue();
        }

        [TestMethod]
        public void ExcelCellStyleFont_TwoInstances_Identical()
        {
            var cellStyleA = new ExcelCellStyleFont { Color = Color.Red, FontName = "test", IsBold = true, IsItalic = false, Size = 99 };
            var cellStyleB = new ExcelCellStyleFont { Color = Color.Red, FontName = "test", IsBold = true, IsItalic = false, Size = 99 };

            cellStyleA.Equals(cellStyleB).ShouldBeTrue();
        }

        [TestMethod]
        public void ExcelCellStyleFont_ListOfTwoInstancesDistinct_ShouldReturnOneInstance()
        {
            var lst = new List<ExcelCellStyleFont> {
                new ExcelCellStyleFont { Color = Color.Red, FontName = "test", IsBold = true, IsItalic = false, Size = 99, IsUnderline= true },
                new ExcelCellStyleFont { Color = Color.Red, FontName = "test", IsBold = true, IsItalic = false, Size = 99, IsUnderline= true }
            };

            lst[0].ShouldBe(lst[1]);
            lst[0].GetHashCode().ShouldBe(lst[1].GetHashCode());
            lst.Distinct().Count().ShouldBe(1);
        }

        [TestMethod]
        public void ExcelCellStyleFont_ListOfThreeInstancesDistinct_ShouldReturnOneInstance()
        {
            var lst = new List<ExcelCellStyleFont> {
                new ExcelCellStyleFont { Color = Color.Red, FontName = "test", IsBold = true, IsItalic = false, Size = 99, IsUnderline= true },
                new ExcelCellStyleFont { Color = Color.Red, FontName = "test", IsBold = true, IsItalic = false, Size = 99, IsUnderline= true },
                new ExcelCellStyleFont { Color = Color.Red, FontName = "test", IsBold = true, IsItalic = false, Size = 99, IsUnderline= true },
            };

            lst.Distinct().Count().ShouldBe(1);
        }

        [TestMethod]
        public void ExcelCellStyleFont_ListOfThreeDifferentInstancesDistinct_ShouldReturnTwoInstances()
        {
            var lst = new List<ExcelCellStyleFont> {
                new ExcelCellStyleFont { Color = Color.Red, FontName = "test", IsBold = true, IsItalic = false, Size = 99, IsUnderline= true },
                new ExcelCellStyleFont { Color = Color.Green, FontName = "test", IsBold = true, IsItalic = false, Size = 99, IsUnderline= true },
                new ExcelCellStyleFont { Color = Color.Red, FontName = "test", IsBold = true, IsItalic = false, Size = 99, IsUnderline= true },
            };

            lst.Distinct().Count().ShouldBe(2);
        }
    }
}
