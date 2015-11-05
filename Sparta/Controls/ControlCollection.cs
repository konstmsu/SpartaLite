using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Sparta.Engine.Utils;

namespace Sparta.Controls
{
    public class ControlCollection
    {
        readonly List<IControl> _controls = new List<IControl>();

        public int Count => _controls.Count;
        public IControl this[int i] => _controls[i];

        internal void OnBeforeDoubleClick(Range target, HandledIndicator handled)
        {
            foreach (var control in _controls)
            {
                var range = control.NarrowDownEventRange(target);

                if (range != null)
                    control.BeforeDoubleClick(range, handled);
            }
        }

        internal void OnChange(Range target)
        {
            _controls.ForEachAggregatingExceptions(control =>
            {
                var range = control.NarrowDownEventRange(target);

                if (range != null)
                    control.OnChange(range);
            });
        }

        internal void Paint()
        {
            foreach (var control in _controls)
                control.Paint();
        }

        internal void Add(IControl button)
        {
            _controls.Add(button);
        }
    }
}
