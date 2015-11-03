using System;
using Microsoft.Office.Interop.Excel;
using Sparta.Controls;
using Sparta.Engine.Utils;

namespace Sparta.Sheets
{
    public class PricerSheet : SheetBase
    {
        public PricerSheet(Worksheet sheet)
            : base(sheet)
        {
            var marketSettings = ControlRoot.AddControl(new PropertyGridControl(sheet.Range["B3"]));

            var market = new DropDownSelector { Values = new[] { "Live", "Close" }.ToReadOnly() };
            var valuationDate = new DateEditorControl { Value = DateTime.Today };
            System.Action onMarketChanged = () => valuationDate.IsDisabled = market.SelectedValue == "Live";
            market.SelectedValueChanged += onMarketChanged;
            onMarketChanged();

            marketSettings.AddProperty("Valuation Date", valuationDate);
            marketSettings.AddProperty("Market", market);
        }
    }
}
