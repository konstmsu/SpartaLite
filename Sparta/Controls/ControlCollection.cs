using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace Sparta.Controls
{
    class ControlCollection
    {
        readonly List<IControl> _controls = new List<IControl>();

        internal void OnBeforeDoubleClick(Range target, ref bool cancel)
        {
            foreach(var control in _controls)
                control.BeforeDoubleClick(target, ref cancel);
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
