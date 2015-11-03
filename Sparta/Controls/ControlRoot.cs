using Microsoft.Office.Interop.Excel;
using System;

namespace Sparta.Controls
{
    public class ControlRoot : IDisposable
    {
        readonly Worksheet _sheet;
        readonly ControlCollection _children = new ControlCollection();

        public ControlRoot(Worksheet sheet)
        {
            _sheet = sheet;

            sheet.Change += OnChange;
            sheet.BeforeDoubleClick += OnBeforeDoubleClick;
        }

        internal void AddControl(IControl control)
        {
            _children.Add(control);
        }

        public void Dispose()
        {
            _sheet.BeforeDoubleClick -= OnBeforeDoubleClick;
            _sheet.Change -= OnChange;
        }

        void OnChange(Range Target)
        {
            _children.OnChange(Target);
        }

        void OnBeforeDoubleClick(Range Target, ref bool Cancel)
        {
            var handled = new HandledIndicator();

            _children.OnBeforeDoubleClick(Target, handled);

            if (handled.IsHandled)
                Cancel = true;

            Paint();
        }

        public void Paint()
        {
            _children.Paint();
        }
    }
}
