using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Sparta.Engine.Utils
{
    public static class ExtensionMethods
    {
        public static string JoinStrings(this IEnumerable<string> values, string separator) => values == null ? null : string.Join(separator, values);
        public static string FormatWith(this string format, params object[] args) => format == null ? null : string.Format(format, args);
        public static string[] Split(this string value, string separator) => value == null ? null : value.Split(new[] { separator }, StringSplitOptions.None);
        public static ReadOnlyCollection<T> ToReadOnly<T>(this IEnumerable<T> values) => values == null ? new List<T>().AsReadOnly() : values.ToList().AsReadOnly();

        public static void ForEachAggregatingExceptions<T>(this IEnumerable<T> values, Action<T> action)
        {
            var exceptions = new List<Exception>();

            foreach (var value in values)
                try
                {
                    action(value);
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                }

            if (exceptions.Any())
                throw new AggregateException(exceptions);
        }
    }

    public static class ReadOnlyCollectionEx<T>
    {
        public static readonly ReadOnlyCollection<T> Empty = new ReadOnlyCollection<T>(new T[0]);
    }
}
