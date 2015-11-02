using Microsoft.Office.Interop.Excel;
using Sparta.Engine.Utils;

namespace Sparta.Controls
{
    public class Button : IControl
    {
        public readonly RangePainter Painter;
        readonly Range _anchor;
        Range Range => _anchor.Resize[2, 2];

        public event System.Action Clicked;

        public string Title
        {
            get { return (string)Painter.Value2Property.Value2; }
            set { Painter.Value2Property.Value2 = value; }
        }

        public Button(Range anchor)
        {
            _anchor = anchor;
            Painter = new RangePainter
            {
                MergeCells = true,
                VerticalAlignment = XlVAlign.xlVAlignCenter,
                HorizontalAlignment = XlHAlign.xlHAlignCenter,
                InteriorColor = SpartaColors.ButtonBackground,
                FontColor = SpartaColors.ButtonForeground,
                IsBold = true,
            };
            Painter.Border.Around();
        }

        public void Paint()
        {
            Painter.Paint(Range);
        }

        public void BeforeDoubleClick(Range target, ref bool handled)
        {
            if (target.Application.Intersect(target, Range) != null)
            {
                Clicked.Raise();
                handled = true;
            }
        }
    }
}
