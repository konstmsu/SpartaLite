using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Sparta.Utils;

namespace Sparta.Controls
{
    public class DataGridControl : IControl
    {
        readonly Func<int> _getRowCount;
        readonly RangePainter _headerPainter = new RangePainter();
        readonly RangePainter _bodyPainter = new RangePainter();

        Range _headerRange;
        Range _bodyRange;

        public DataGridControl(Range anchor, Func<int> getRowCount)
        {
            _getRowCount = getRowCount;
            Anchor = anchor;
        }

        class DataGridColumn
        {
            public readonly IControl Header;
            public Func<int, IControl> GetCell;

            public DataGridColumn(IControl header, Func<int, IControl> getCell)
            {
                Header = header;
                GetCell = getCell;
            }
        }

        readonly List<DataGridColumn> _columns = new List<DataGridColumn>();

        public void AddColumn(string header, Func<int, IControl> getCell)
        {
            _columns.Add(new DataGridColumn(new LabelControl { Text = header }, getCell));
        }

        int ColumnCount => _columns.Count;
        int BodyRowCount => _getRowCount();

        public void Paint()
        {
            _headerRange = Anchor.Resize[1, ColumnCount];
            _bodyRange = Anchor.Offset[1].Resize[BodyRowCount, ColumnCount];

            _headerPainter.Paint(_headerRange);
            _bodyPainter.Paint(_bodyRange);

            for (var c = 0; c < ColumnCount; c++)
            {
                _columns[c].Header.Anchor = Anchor.Offset[0, c];
                _columns[c].Header.Paint();

            }

            for (var c = 0; c < ColumnCount; c++)
                for (var r = 0; r < BodyRowCount; r++)
                {
                    var cell = _columns[c].GetCell(r);
                    cell.Anchor = Anchor.Offset[r + 1, c];
                    cell.Paint();
                }
        }

        public void BeforeDoubleClick(Range target, HandledIndicator handled)
        {
            new[] { GetCell(target) }.OnBeforeDoubleClick(target, handled);
        }

        IControl GetCell(Range target)
        {
            var row = target.Row - Anchor.Row - 1;
            var column = target.Column - Anchor.Column;

            return row == -1 ? _columns[column].Header : _columns[column].GetCell(row);
        }

        public Range Anchor { get; set; }
        public Range NarrowDownEventRange(Range target)
        {
            return target.GetIntersection(_headerRange) ?? target.GetIntersection(_bodyRange);
        }

        public void OnChange(Range target)
        {
            target.Cells.Cast<Range>().Select(GetCell).OnChange(target);

        }
    }
}
