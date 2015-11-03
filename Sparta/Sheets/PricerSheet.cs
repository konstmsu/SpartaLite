using System;
using Microsoft.Office.Interop.Excel;
using Sparta.Controls;

namespace Sparta.Sheets
{
    public class PricerSheet : SheetBase
    {
        public PricerSheet(Worksheet sheet)
            : base(sheet)
        {
            PropertyGridControl market;
            ControlRoot.AddControl(market = new PropertyGridControl(sheet.Range["B3"]));
            market.AddProperty("Valuation Date", new DateEditorControl { Value = DateTime.Today });
            market.AddProperty("Market", new LabelControl { Text = "Live" });
        }
    }
}
