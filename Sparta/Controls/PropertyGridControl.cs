using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Sparta.Controls
{
    public class PropertyGridControl : IControl
    {
        public Range Anchor { get; set; }
        public Range NarrowDownEventRange(Range target)
        {
            return target;
        }

        readonly ControlCollection _labels = new ControlCollection();
        readonly ControlCollection _values = new ControlCollection();

        public PropertyGridControl(Range anchor)
        {
            Anchor = anchor;
        }

        public void AddProperty(string title, IControl value)
        {
            _labels.Add(new LabelControl { Text = title });
            _values.Add(value);
        }

        public void Paint()
        {
            var count = _labels.Count;
            Debug.Assert(_values.Count == count);

            for (var i = 0; i < count; i++)
            {
                _labels[i].Anchor = Anchor.Offset[i];
                _values[i].Anchor = Anchor.Offset[i, 1];
            }

            // TODO: Optimize painting
            _labels.Paint();
            _values.Paint();
        }

        public void BeforeDoubleClick(Range target, HandledIndicator handled)
        {
            _labels.OnBeforeDoubleClick(target, handled);
            _values.OnBeforeDoubleClick(target, handled);
        }
    }
}
