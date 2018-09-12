using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;

namespace OpenXML.ExcelWrapper.Tests
{
    [TestClass]
    public class ExcelCell_ColumnLetters
    {
        [TestMethod]
        public void GetColumnLetters_OneLetter_A()
        {
            var letters = ExcelCell.GetColumnLetters(1);
            letters.ShouldBe("A");
        }

        [TestMethod]
        public void GetColumnLetters_OneLetter_C()
        {
            var letters = ExcelCell.GetColumnLetters(3);
            letters.ShouldBe("C");
        }

        [TestMethod]
        public void GetColumnLetters_TwoLetters_IV()
        {
            var letters = ExcelCell.GetColumnLetters(256);
            letters.ShouldBe("IV");
        }

        [TestMethod]
        public void GetColumnLetters_TwoLetters_PC()
        {
            var letters = ExcelCell.GetColumnLetters(419);
            letters.ShouldBe("PC");
        }

        [TestMethod]
        public void GetColumnLetters_ThreeLetters_BDJ()
        {
            var letters = ExcelCell.GetColumnLetters(1466);
            letters.ShouldBe("BDJ");
        }
    }
}
