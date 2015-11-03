using Microsoft.Office.Interop.Excel;
using Sparta.Engine.Utils;
using Sparta.Utils;

namespace Sparta.Controls
{
    public class Button : IControl
    {
        public readonly RangePainter Painter;
        public Range Anchor { get; set; }
        public Range NarrowDownEventRange(Range target)
        {
            return target.GetIntersection(Range);
        }

        Range Range => Anchor.Resize[2, 2];

        public event System.Action Clicked;

        public string Title
        {
            get { return (string)Painter.Value2Property.Value2; }
            set { Painter.Value2Property.Value2 = value; }
        }

        public Button(Range anchor)
        {
            Anchor = anchor;
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

        public void BeforeDoubleClick(Range target, HandledIndicator handled)
        {
            if (target.Application.Intersect(target, Range) != null)
            {
                Clicked.Raise();
                handled.MarkHandled();
            }
        }
    }
}
