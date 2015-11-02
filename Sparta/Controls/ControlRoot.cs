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

        internal void AddControl(Button button)
        {
            _children.Add(button);
        }

        public void Dispose()
        {
            _sheet.BeforeDoubleClick -= OnBeforeDoubleClick;
            _sheet.Change -= OnChange;
        }

        private void OnChange(Range Target)
        {
            _children.OnChange(Target);
        }

        private void OnBeforeDoubleClick(Range Target, ref bool Cancel)
        {
            _children.OnBeforeDoubleClick(Target, ref Cancel);
        }

        public void Paint()
        {
            _children.Paint();
        }
    }
}
