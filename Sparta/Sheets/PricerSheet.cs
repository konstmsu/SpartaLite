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

            var market = new DropDownSelector { Values = new[] { "Live", "Close" }.ToReadOnly(), SelectedValue = "Live" };
            var valuationDate = new DateEditorControl { Value = Today };
            System.Action onMarketChanged = () => valuationDate.IsDisabled = market.SelectedValue == "Live";
            market.SelectedValueChanged += onMarketChanged;
            onMarketChanged();

            marketSettings.AddProperty("Valuation Date", valuationDate);
            marketSettings.AddProperty("Market", market);

            var trades = new List<TradeRowView>();

            trades.Add(new FXForwardRowView
            {
                Domestic = { Value = Eur(10000) },
                Foreign = { Value = Usd(12000) },
                SettlementDate = { Value = Today.AddMonths(3) },
                ForwardRate = { Value = 1.2m }
            });

            trades.Add(new FXVanillaOptionRowView
            {
                Domestic = { Value = Rub(10000) },
                Foreign = { Value = Sgd(-12000) },
                SettlementDate = { Value = Today.AddMonths(6) },
                Strike = { Value = 1.4m }
            });

            var tradeArea = ControlRoot.AddControl(new DynamicDataGridControl(sheet.Range["B15"], trades));

            tradeArea.AddColumns(new[] { "Settlement Date", "Domestic" });
        }
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
        public readonly DecimalEditorControl ForwardRate;

        public FXForwardRowView()
        {
            AddProperty("Settlement Date", SettlementDate = new DateEditorControl());
            AddProperty("Domestic", Domestic = new MoneyControl());
            AddProperty("Foreign", Foreign = new MoneyControl());
            AddProperty("Forward Rate", ForwardRate = new DecimalEditorControl());
        }
    }

    public class FXVanillaOptionRowView : TradeRowView
    {
        public readonly DateEditorControl SettlementDate;
        public readonly MoneyControl Domestic;
        public readonly MoneyControl Foreign;
        public readonly DecimalEditorControl Strike;

        public FXVanillaOptionRowView()
        {
            AddProperty("Settlement Date", SettlementDate = new DateEditorControl());
            AddProperty("Domestic", Domestic = new MoneyControl());
            AddProperty("Foreign", Foreign = new MoneyControl());
            AddProperty("Strike", Strike = new DecimalEditorControl());
        }
    }
}
