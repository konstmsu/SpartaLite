using Microsoft.Office.Interop.Excel;
using Sparta.Utils;

namespace Sparta.Controls
{
    public class LabelControl : IControl
    {
        public Range Anchor { get; set; }
        readonly int _rowCount;
        readonly int _columnCount;
        Range Range => Anchor.Resize[_rowCount, _columnCount];

        public Range NarrowDownEventRange(Range target)
        {
            return target.GetIntersection(Range);
        }

        public void OnChange(Range target)
        {
        }

        public readonly RangePainter Painter;

        public string Text
        {
            get { return (string)Painter.Value2; }
            set { Painter.Value2 = value; }
        }

        public LabelControl(int rowCount = 1, int columnCount = 1)
        {
            _rowCount = rowCount;
            _columnCount = columnCount;

            Painter = new RangePainter
            {
                InteriorColor = ExcelColor.LightGray,
                MergeCells = true
            };
        }

        public void Paint()
        {
            Painter.Paint(Range);
        }

        public void BeforeDoubleClick(Range target, HandledIndicator handled)
        {
            // Nothing should happen here
        }
    }
}
