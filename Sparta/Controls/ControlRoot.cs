using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using Sparta.Utils;

namespace Sparta.Controls
{
    public class ControlRoot : IDisposable
    {
        readonly Worksheet _sheet;
        readonly List<IControl> _children = new List<IControl>();
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
            using (_sheet.Application.SetTemporarily(enableEvents: false, screenUpdating: false))
            {
                _children.Paint();
                _sheet.Columns.AutoFit();
            }
        }
    }
}
