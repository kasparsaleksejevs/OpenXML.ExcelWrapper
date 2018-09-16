using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXML.ExcelWrapper.Styling;
using Shouldly;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace OpenXML.ExcelWrapper.Tests
{
    [TestClass]
    public class ExcelCellStyleBorder_Equality
    {
        [TestMethod]
        public void ExcelCellStyleBorder_OneInstance()
        {
            var cellStyleA = new ExcelCellStyleBorder(ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Dashed, new ExcelColor("559999"));
            var cellStyleB = cellStyleA;
            cellStyleA.Equals(cellStyleB).ShouldBeTrue();
        }

        [TestMethod]
        public void ExcelCellStyleBorder_TwoInstances_Identical()
        {
            var cellStyleA = new ExcelCellStyleBorder(ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Dashed, new ExcelColor("559999"));
            var cellStyleB = new ExcelCellStyleBorder(ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Dashed, new ExcelColor("559999"));

            cellStyleA.Equals(cellStyleB).ShouldBeTrue();
        }

        [TestMethod]
        public void ExcelCellStyleBorder_ListOfTwoInstancesDistinct_ShouldReturnOneInstance()
        {
            var lst = new List<ExcelCellStyleBorder> {
                new ExcelCellStyleBorder(ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Dashed, new ExcelColor("559999")),
                new ExcelCellStyleBorder(ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Dashed, new ExcelColor("559999"))
            };

            lst[0].ShouldBe(lst[1]);
            lst[0].GetHashCode().ShouldBe(lst[1].GetHashCode());
            lst.Distinct().Count().ShouldBe(1);
        }

        [TestMethod]
        public void ExcelCellStyleBorder_ListOfThreeInstancesDistinct_ShouldReturnOneInstance()
        {
            var lst = new List<ExcelCellStyleBorder> {
                new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.Double, new ExcelColor("00FF00")),
                new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.Double, new ExcelColor("00FF00")),
                new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.Double, new ExcelColor("00FF00")),
            };

            lst.Distinct().Count().ShouldBe(1);
        }

        [TestMethod]
        public void ExcelCellStyleBorder_ListOfThreeDifferentInstancesDistinct_ShouldReturnTwoInstances()
        {
            var lst = new List<ExcelCellStyleBorder> {
                new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.Double, new ExcelColor("00FF00")),
                new ExcelCellStyleBorder(ExcelCellBorderEnum.Bottom, ExcelCellStyleBorderSizeEnum.Dashed, new ExcelColor("559999")),
                new ExcelCellStyleBorder(ExcelCellBorderEnum.Top, ExcelCellStyleBorderSizeEnum.Double, new ExcelColor("00FF00")),
            };

            lst.Distinct().Count().ShouldBe(2);
        }
    }
}
