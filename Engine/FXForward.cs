using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using static Sparta.Engine.Currency;

namespace Sparta.Engine
{
    public class FXForward
    {
        public DateTime SettlementDate;
        public Money Domestic;
        public Money Foreign;
    }

    public class Money : IEquatable<Money>
    {
        public readonly Currency Currency;
        public readonly decimal Amount;

        public Money(Currency currency, decimal amount)
        {
            Currency = currency;
            Amount = amount;
        }

        public static Money Eur(decimal amount) => new Money(EUR, amount);
        public static Money Usd(decimal amount) => new Money(USD, amount);
        public static Money Sgd(decimal amount) => new Money(SGD, amount);
        public static Money Rub(decimal amount) => new Money(RUB, amount);

        public static bool TryParse(string input, Currency defaultCurrency, out Money result)
        {
            var formats = new[]
            {
                @"^(?<amount>-?\d+)(?<multiplier>n?[km]?n?)\s*(?<currency>[a-z]+)?$",
                @"^(?<currency>[a-z]+)\s*(?<amount>-?\d+)(?<multiplier>n?[km]?n?)$",
            };

            var matched = formats.Select(f =>
            {
                var format1 = Regex.Match(input, f, RegexOptions.IgnoreCase);

                if (!format1.Success)
                    return null;

                var currencyStr = format1.Groups["currency"].Value;
                Currency currency;
                if (string.IsNullOrEmpty(currencyStr))
                    currency = defaultCurrency;
                else
                {
                    if (!Currency.TryParse(currencyStr, out currency))
                        return null;
                }
                var amount = decimal.Parse(format1.Groups["amount"].Value);
                var multiplier = format1.Groups["multiplier"].Value.ToUpperInvariant();

                var withoutNegative = multiplier.Replace("N", "");

                if (withoutNegative != multiplier)
                {
                    // Double negative probably means user error
                    if (amount < 0)
                        return null;

                    amount *= -1;
                }

                if (withoutNegative == "K")
                    amount *= 1000;
                else if (withoutNegative == "M")
                    amount *= 1000 * 1000;

                return new Money(currency, amount);
            }).FirstOrDefault(v => v != null);

            result = matched;
            return matched != null;
        }

        public bool Equals(Money other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(Currency, other.Currency) && Amount == other.Amount;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((Money)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((Currency != null ? Currency.GetHashCode() : 0) * 397) ^ Amount.GetHashCode();
            }
        }

        public static bool operator ==(Money left, Money right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(Money left, Money right)
        {
            return !Equals(left, right);
        }
    }

    public class Currency
    {
        public readonly string Code;

        Currency(string code)
        {
            Code = code;
        }

        public static readonly Currency USD = new Currency("USD");
        public static readonly Currency EUR = new Currency("EUR");
        public static readonly Currency SGD = new Currency("SGD");
        public static readonly Currency RUB = new Currency("RUB");

        static IEnumerable<Currency> KnownCurrencies
        {
            get
            {
                yield return EUR;
                yield return RUB;
                yield return SGD;
                yield return USD;
            }
        }

        public static bool TryParse(string input, out Currency result)
        {
            result =  KnownCurrencies.FirstOrDefault(c => c.Code.Equals(input, StringComparison.OrdinalIgnoreCase));
            return result != null;
        }
    }
}
