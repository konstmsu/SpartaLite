using System;
using Microsoft.Office.Interop.Excel;
using Sparta.Engine;
using Sparta.Utils;

namespace Sparta.Controls
{
    public class MoneyControl : IControl
    {
        Money _value;

        public Money Value
        {
            get { return _value; }
            set
            {
                _value = value;
                _painter.Value2 = value.Amount;
                var currency = $"\"{value.Currency.Code}\"";
                _painter.NumberFormat = $"{currency} #,0.00;[Red]{currency} (#,0.00)";
            }
        }

        readonly RangePainter _painter = new RangePainter();

        public Range Anchor { get; set; }

        public void Paint()
        {
            _painter.Paint(Anchor);
        }

        public void BeforeDoubleClick(Range target, HandledIndicator handled)
        {
            throw new NotImplementedException();
        }

        public Range NarrowDownEventRange(Range target)
        {
            return target.GetIntersection(Anchor);
        }

        public void OnChange(Range target)
        {
            Money result;

            if (target.Value2 is decimal)
                Value = new Money(Value.Currency, (decimal)target.Value2);
            else if (target.Value2 is double)
                Value = new Money(Value.Currency, (decimal)(double)target.Value2);
            else if (target.Value2 is string && Money.TryParse((string)target.Value2, Value.Currency, out result))
                Value = result;
            else
                throw new FormatException($"Could not parse '{target.Value2}' as Money");
        }
    }
}