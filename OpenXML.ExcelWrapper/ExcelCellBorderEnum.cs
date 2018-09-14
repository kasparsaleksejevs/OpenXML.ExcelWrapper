using System;

namespace OpenXML.ExcelWrapper
{
    [Flags]
    public enum ExcelCellBorderEnum
    {
        None = 0,
        Left = 1,
        Right = 2,
        Top = 4,
        Bottom = 8,
        Diagonal = 16
    }
}
