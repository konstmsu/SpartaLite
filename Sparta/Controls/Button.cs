using Microsoft.Office.Interop.Excel;

namespace Sparta.Controls
{
    public class Button : IControl
    {
        readonly RangePainter _painter;
        readonly Range _anchor;

        public string Title
        {
            get { return (string)_painter.Value2Property.Value2; }
            set { _painter.Value2Property.Value2 = value; }
        }

        public Button(Range anchor)
        {
            _anchor = anchor;
            _painter = new RangePainter(anchor.Worksheet)
            {
                MergeCells = true,
                VerticalAlignment = XlVAlign.xlVAlignCenter,
                HorizontalAlignment = XlHAlign.xlHAlignCenter,
                InteriorColor = SpartaColors.ButtonBackground,
            };
        }

        public void Paint()
        {
            _painter.Paint(_anchor.Resize[2, 2]);
        }
    }
}
