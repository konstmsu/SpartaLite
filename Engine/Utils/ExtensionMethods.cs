using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Sparta.Engine.Utils
{
    public static class ExtensionMethods
    {
        public static void Raise(this Action action) => action?.Invoke();
        public static string JoinStrings(this IEnumerable<string> values, string separator)
        {
            return values == null ? null : string.Join(separator, values);
        }

        public static string FormatWith(this string format, params object[] args) => string.Format(format, args);

        public static ReadOnlyCollection<T> ToReadOnly<T>(this IEnumerable<T> values) => values.ToList().AsReadOnly();
    }

    public static class ReadOnlyCollectionEx<T>
    {
        public static readonly ReadOnlyCollection<T> Empty = new ReadOnlyCollection<T>(new T[0]);
    }
}
