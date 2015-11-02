using Microsoft.Office.Interop.Excel;
using Sparta.Controls;
using Button = Sparta.Controls.Button;

namespace Sparta.Sheets
{
    public class ContentsSheet : SheetBase
    {
        readonly Button _pricer;

        public ContentsSheet(Worksheet sheet)
            : base(sheet)
        {
            _controlRoot.AddControl(_pricer = new Button(sheet.Range["B3"])
            {
                Title = "Pricer",
            });
            _pricer.Clicked += () => { _pricer.Title += "a"; };
        }

        internal void Run()
        {
            _controlRoot.Paint();
        }
    }
}
