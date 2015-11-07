using System;
using System.Collections.Generic;
using System.Linq;
using Sparta.Engine.Utils;

namespace Sparta.Utils
{
    public class CompositeDisposable : IDisposable
    {
        readonly IReadOnlyCollection<IDisposable> _children;

        public CompositeDisposable(IReadOnlyCollection<IDisposable> children)
        {
            _children = children;
        }

        public void Dispose()
        {
            _children.Reverse().ForEachAggregatingExceptions(c => c.Dispose());
        }
    }
}