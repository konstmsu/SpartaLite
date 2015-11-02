using Microsoft.Office.Interop.Excel;

namespace Sparta.Controls
{
    public class BorderProperty : IRangeProperty
    {
        bool _isBorderAround;
        ExcelColor _color;

        public void Around(ExcelColor color = null)
        {
            _isBorderAround = true;
            _color = color;
        }

        public void Paint(Range range)
        {
            if (_isBorderAround)
                range.BorderAround(Color: _color?.Code);
        }
    }
}