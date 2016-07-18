using System;
using NUnit.Framework;
using EchangeExporterProto;

namespace ExchangeExporter.Tests
{
    public class PartialExtensionMethod
    {
        [Test]
        public void PassProvidedValueAsFirstParameter()
        {
            Func<int, int, int> add = (x, y) => x;

            var firstParameter = 5;
            var sut = add.Partial(firstParameter);

            var actual = sut.Invoke(0);
            Assert.AreEqual(firstParameter, actual);
        }
    }

    public class MemoizeExtensionMethod
    {
        [Test]
        public void CachesFirstCallResults()
        {
            Func<int, long> expensiveComputations = _ => DateTime.UtcNow.Ticks;
            var sut = expensiveComputations.Memoize();

            var firstResult = sut(0);
            for (int i = 0; i < 10; i++)
            {
                var actual = sut(0);
                Assert.AreEqual(firstResult, actual, "memoized function should return cached results");
            }
        }
    }
}