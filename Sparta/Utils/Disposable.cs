using System;

namespace Sparta.Utils
{
    public class Disposable : IDisposable
    {
        readonly Action _disposed;

        public Disposable(Action disposed)
        {
            _disposed = disposed;
        }

        public void Dispose() => _disposed?.Invoke();

        public static readonly Disposable Empty = new Disposable(null);
    }
}