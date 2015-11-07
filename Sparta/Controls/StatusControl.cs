using System;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Sparta.Engine.Utils;
using Sparta.Utils;

namespace Sparta.Controls
{
    public class StatusControl : IControl
    {
        readonly RangePainter _headerPainter;
        Range HeaderRange => Anchor.Resize[_headerRowCount, _columnCount];

        readonly LabelControl _header;
        readonly LabelControl _body;

        readonly int _columnCount;
        readonly int _headerRowCount;

        public StatusControl(Range anchor, int columnCount, int headerRowCount, int bodyRowCount)
        {
            _columnCount = columnCount;
            _headerRowCount = headerRowCount;

            _headerPainter = new RangePainter
            {
                InteriorColor = SpartaColors.DisabledControlInterior,
            };
            _headerPainter.Border.Around();

            _header = new LabelControl
            {
                Text = "Status",
                Anchor = anchor,
                Painter = { HorizontalAlignment = XlHAlign.xlHAlignLeft, FontSize = 16 }
            };

            _body = new LabelControl(bodyRowCount, columnCount)
            {
                Anchor = anchor.Offset[headerRowCount],
                Painter =
                {
                    VerticalAlignment = XlVAlign.xlVAlignTop,
                    HorizontalAlignment = XlHAlign.xlHAlignLeft,
                }
            };
            _body.Painter.Border.Around();

            Anchor = anchor;
        }

        public Range Anchor { get; set; }

        public void Paint()
        {
            _headerPainter.Paint(HeaderRange);
            new[] { _header, _body }.Paint();
        }

        public void BeforeDoubleClick(Range target, HandledIndicator handled)
        {
            new[] { _header, _body }.OnBeforeDoubleClick(target, handled);
        }

        public Range NarrowDownEventRange(Range target)
        {
            return target;
        }

        public void OnChange(Range target)
        {
            if (_body.NarrowDownEventRange(target) != null)
                _body.Text = null;
        }

        public void Append(Exception exception)
        {
            _body.Text += FormatExceptionMessage(exception);
        }

        static string FormatExceptionMessage(Exception exception)
        {
            if (exception == null)
                return null;

            var inners = CoerceTo<string[]>.Value(exception)
                .Type<AggregateException>(a => a.InnerExceptions.Select(FormatExceptionMessage).ToArray())
                .Else(e => new[] { FormatExceptionMessage(e.InnerException) });

            var aggregate = exception as AggregateException;

            if (aggregate != null)
                if (aggregate.Message == "One or more errors occurred.")
                    return inners.JoinStrings(Environment.NewLine);

            return exception.Message + Environment.NewLine + inners.Select(i => Indent(i)).JoinStrings(Environment.NewLine);
        }

        static string Indent(string value, int size = 1)
        {
            if (value == null)
                return null;

            var indent = new string(' ', 4 * size);
            return value.Split(Environment.NewLine).Select(l => string.IsNullOrWhiteSpace(l) ? l : indent + l).JoinStrings(Environment.NewLine);
        }
    }
}