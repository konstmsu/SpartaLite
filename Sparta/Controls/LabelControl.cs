using Microsoft.Office.Interop.Excel;
using Sparta.Utils;

namespace Sparta.Controls
{
    public class LabelControl : IControl
    {
        public Range Anchor { get; set; }
        public Range NarrowDownEventRange(Range target)
        {
            return target.GetIntersection(Anchor);
        }

        public void OnChange(Range target)
        {
        }

        readonly RangePainter _painter;

        public string Text
        {
            get { return (string)_painter.Value2; }
            set { _painter.Value2 = value; }
        }

        public LabelControl()
        {
            _painter = new RangePainter
            {
                InteriorColor = ExcelColor.LightGray
            };
        }

        public void Paint()
        {
            _painter.Paint(Anchor);
        }

        public void BeforeDoubleClick(Range target, HandledIndicator handled)
        {
            // Nothing should happen here
        }
    }
}
