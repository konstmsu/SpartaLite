using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Sparta.Controls;
using Sparta.Engine.Utils;
using static System.DateTime;
using static Sparta.Engine.Money;

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

            var tradeArea = ControlRoot.AddControl(new DataGridControl(sheet.Range["B15"], () => _trades.Count));
            tradeArea.AddColumn("Settlement Date", i => _trades[i].SettlementDate);
            tradeArea.AddColumn("Domestic", i => _trades[i].Domestic);
            tradeArea.AddColumn("Foreign", i => _trades[i].Foreign);

            _trades.Add(new FXForwardRowView
            {
                Domestic = { Value = Eur(10000) },
                Foreign = { Value = Usd(12000) },
                SettlementDate = { Value = Today.AddMonths(3) }
            });

            _trades.Add(new FXForwardRowView
            {
                Domestic = { Value = Rub(10000) },
                Foreign = { Value = Sgd(-12000) },
                SettlementDate = { Value = Today.AddMonths(6) }
            });
        }

        readonly List<FXForwardRowView> _trades = new List<FXForwardRowView>();
    }

    public class TradePropertyView
    {
        public readonly string Header;
        public readonly IControl Control;

        public TradePropertyView(string header, IControl control)
        {
            Header = header;
            Control = control;
        }
    }

    public class TradeRowView
    {
        public List<TradePropertyView> Properties = new List<TradePropertyView>();

        protected void AddProperty(string header, IControl control)
        {
            Properties.Add(new TradePropertyView(header, control));
        }
    }

    public class FXForwardRowView : TradeRowView
    {
        public readonly DateEditorControl SettlementDate;
        public readonly MoneyControl Domestic;
        public readonly MoneyControl Foreign;

        public FXForwardRowView()
        {
            AddProperty("Settlement Date", SettlementDate = new DateEditorControl());
            AddProperty("Domestic", Domestic = new MoneyControl());
            AddProperty("Foreign", Foreign = new MoneyControl());
        }
    }
}
