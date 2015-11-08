using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Sparta.Engine.Utils;
using Sparta.Sheets;
using Sparta.Utils;

namespace Sparta.Controls
{
    public class DynamicDataGridControl : IControl
    {
        readonly IReadOnlyList<TradeRowView> _rows;
        readonly List<DataGridColumn> _columns = new List<DataGridColumn>();
        readonly RangePainter _headerPainter = new RangePainter();
        readonly RangePainter _bodyPainter = new RangePainter();

        Range _headerRange;
        Range _bodyRange;

        public DynamicDataGridControl(Range anchor, IReadOnlyList<TradeRowView> rows)
        {
            _rows = rows;
            Anchor = anchor;
        }

        class DataGridColumn
        {
            public readonly DropDownSelector Header;

            public DataGridColumn(string title, Func<ReadOnlyCollection<string>> getAllHeaders)
            {
                Header = new DropDownSelector { SelectedValue = title, Values = getAllHeaders() };
            }

            public TradePropertyView GetCell(TradeRowView row)
            {
                return row.Properties.SingleOrDefault(p => p.Header == Header.SelectedValue);
            }
        }

        void AddColumn(string title)
        {
            _columns.Add(new DataGridColumn(title, GetAllHeaders));
        }

        ReadOnlyCollection<string> GetAllHeaders()
        {
            return _rows.SelectMany(r => r.Properties).Select(p => p.Header).Distinct().ToReadOnly();
        }

        int ColumnCount => _columns.Count;
        int BodyRowCount => _rows.Count;

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
                    var property = _columns[c].GetCell(_rows[r]);

                    if (property == null)
                        continue;

                    var cell = property.Control;
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

            return row == -1 ? _columns[column].Header : _columns[column].GetCell(_rows[row])?.Control;
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

        public void AddColumns(IEnumerable<string> columnHeaders)
        {
            foreach (var header in columnHeaders)
                AddColumn(header);
        }
    }
}
