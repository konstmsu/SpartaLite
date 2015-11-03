using Microsoft.Office.Interop.Excel;
using Sparta.Controls;
using Button = Sparta.Controls.Button;

namespace Sparta.Sheets
{
    public class ContentsSheet : SheetBase
    {
        readonly Button _pricer;

        public ContentsSheet(Worksheet sheet, SheetFactory sheetFactory)
            : base(sheet)
        {
            ControlRoot.AddControl(_pricer = new Button(sheet.Range["B3"])
            {
                Title = "Pricer",
            });
            _pricer.Clicked += () => { sheetFactory.ShowSheet(s => new PricerSheet(s)); };
        }
    }
}
