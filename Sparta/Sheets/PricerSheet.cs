using System;
using Microsoft.Office.Interop.Excel;
using Sparta.Controls;
using Sparta.Engine.Utils;

namespace Sparta.Sheets
{
    public class PricerSheet : SheetBase
    {
        readonly StatusControl _status;

        public PricerSheet(Worksheet sheet)
            : base(sheet)
        {
            _status = ControlRoot.AddControl(new StatusControl(sheet.Range["I4"], 10, 1, 6));
            ControlRoot.UnhandledException += ex =>
            {
                _status.Append(ex);
                // TODO: _log.Warn(ex);
            };

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
