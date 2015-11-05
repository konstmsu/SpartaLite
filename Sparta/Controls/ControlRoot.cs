using Microsoft.Office.Interop.Excel;
using System;

namespace Sparta.Controls
{
    public class ControlRoot : IDisposable
    {
        readonly Worksheet _sheet;
        readonly ControlCollection _children = new ControlCollection();
        public event Action<Exception> UnhandledException;

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

        void OnChange(Range target)
        {
            try
            {
                _children.OnChange(target);
            }
            catch (Exception ex)
            {
                UnhandledException?.Invoke(ex);
            }

            Paint();
        }

        void OnBeforeDoubleClick(Range target, ref bool cancel)
        {
            var handled = new HandledIndicator();

            _children.OnBeforeDoubleClick(target, handled);

            if (handled.IsHandled)
                cancel = true;

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
