using Microsoft.Office.Interop.Excel;
using Sparta.Controls;
using Button = Sparta.Controls.Button;

namespace Sparta.Sheets
{
    public class ContentsSheet : SheetBase
    {
        public ContentsSheet(Worksheet sheet, SheetFactory sheetFactory)
            : base(sheet)
        {
            var pricer = ControlRoot.AddControl(new Button(sheet.Range["B3"])
            {
                Title = "Pricer"
            });

            pricer.Clicked += () => sheetFactory.ShowSheet(s => new PricerSheet(s));
        }
    }
}
