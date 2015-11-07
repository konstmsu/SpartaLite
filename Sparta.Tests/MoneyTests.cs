using System;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Sparta.Engine;
using static Sparta.Engine.Currency;
using static Sparta.Engine.Money;

namespace Sparta.Tests
{
    [TestClass]
    public class MoneyTests
    {
        [TestMethod]
        public void ShouldEqual()
        {
            Action<Money, Money, bool> assertOneWay = (left, right, expectedAreEqual) =>
            {
                left.Equals(right).Should().Be(expectedAreEqual);
                (left == right).Should().Be(expectedAreEqual);
                (left != right).Should().Be(!expectedAreEqual);
            };

            Action<Money, Money, bool> assert = (left, right, expectedAreEqual) =>
            {
                assertOneWay(left, right, expectedAreEqual);
                assertOneWay(right, left, expectedAreEqual);
            };

            assert(Eur(100), Eur(100), true);
            assert(Eur(100), Usd(100), false);
            assert(Eur(100), Eur(101), false);
        }

        [TestMethod]
        public void ShouldTryParse()
        {
            Money value;
            TryParse("100", EUR, out value).Should().BeTrue();
            value.Should().Be(Eur(100));

            TryParse("2000", USD, out value).Should().BeTrue();
            value.Should().Be(Usd(2000));

            TryParse("-150", EUR, out value).Should().BeTrue();
            value.Should().Be(Eur(-150));

            TryParse("10k USD", RUB, out value).Should().BeTrue();
            value.Should().Be(Usd(10000));

            TryParse("eur 5nm", EUR, out value).Should().BeTrue();
            value.Should().Be(Eur(-5000000));

            TryParse("rub -3k", EUR, out value).Should().BeTrue();
            value.Should().Be(Rub(-3000));

            TryParse("5e", USD, out value).Should().BeFalse();
            TryParse("abc", USD, out value).Should().BeFalse();
            TryParse("usd", USD, out value).Should().BeFalse();
            TryParse("gbp", USD, out value).Should().BeFalse();
        }
    }
}