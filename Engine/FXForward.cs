using System;

namespace Sparta.Engine
{
    public class FXForward
    {
        public DateTime SettlementDate;
        public Money Domestic;
        public Money Foreign;
    }

    public class Money
    {
        public readonly Currency Currency;
        public readonly decimal Amount;

        public Money(Currency currency, decimal amount)
        {
            Currency = currency;
            Amount = amount;
        }
    }

    public class Currency
    {
        readonly string _code;

        Currency(string code)
        {
            _code = code;
        }

        public static readonly Currency USD = new Currency("USD");
        public static readonly Currency EUR = new Currency("EUR");
        public static readonly Currency SGD = new Currency("SGD");
    }
}
