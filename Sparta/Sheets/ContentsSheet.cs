using Microsoft.Office.Interop.Excel;
using Sparta.Controls;

namespace Sparta.Sheets
{
    public class ContentsSheet : SheetBase
    {
        public ContentsSheet(Worksheet sheet)
            : base(sheet)
        {
            _controlRoot.AddControl(new Controls.Button(sheet.Range["B3"])
            {
                Title = "FX Option Pricer"
            });
        }

        internal void Run()
        {
            _controlRoot.Paint();
        }
    }
}
