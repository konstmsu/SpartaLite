using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace Sparta.Controls
{
    class ControlCollection
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
                    control.BeforeDoubleClick(target, handled);
            }
        }

        internal void OnChange(Range target)
        {
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
