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

        internal T AddControl<T>(T control)
            where T: IControl
        {
            _children.Add(control);
            return control;
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
            // TODO: Preserve old value
            _sheet.Application.EnableEvents = false;
            try
            {
                _children.Paint();
                _sheet.Columns.AutoFit();
            }
            finally
            {
                _sheet.Application.EnableEvents = true;
            }
        }
    }
}
