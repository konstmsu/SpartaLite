using System;
using Sparta.Engine.Utils;

namespace Sparta.Utils
{
    public static class CoerceTo<TTo>
    {
        public static ICoercionContext<TFrom> Value<TFrom>(TFrom value)
        {
            return new CoercionContext<TFrom>(value);
        }

        public interface ICoercionContext<TFrom>
        {
            ICoercionContext<TFrom> Type<T>(Func<T, TTo> coerce);
            ICoercionContext<TFrom> Null(Func<TTo> result);
            ICoercionContext<TFrom> If(Func<TFrom, bool> predicate, Func<TFrom, TTo> result);
            TTo Else(Func<TFrom, TTo> coerce);
            TTo ElseThrow();
        }

        public class CoercionComplete<TFrom> : ICoercionContext<TFrom>
        {
            readonly TTo result;

            public CoercionComplete(TTo result)
            {
                this.result = result;
            }

            public ICoercionContext<TFrom> Null(Func<TTo> value)
            {
                return this;
            }

            public ICoercionContext<TFrom> If(Func<TFrom, bool> predicate, Func<TFrom, TTo> result)
            {
                return this;
            }

            public ICoercionContext<TFrom> Type<T>(Func<T, TTo> coerce)
            {
                return this;
            }

            public TTo Else(Func<TFrom, TTo> coerce)
            {
                return result;
            }

            public TTo ElseThrow()
            {
                return result;
            }
        }

        public class CoercionContext<TFrom> : ICoercionContext<TFrom>
        {
            readonly TFrom value;

            public CoercionContext(TFrom value)
            {
                this.value = value;
            }

            public ICoercionContext<TFrom> Type<T>(Func<T, TTo> coerce)
            {
                if (value is T)
                {
                    // TODO: Get rid of double-cast
                    return Complete(coerce((T)(object)value));
                }

                return this;
            }

            public ICoercionContext<TFrom> If(Func<TFrom, bool> predicate, Func<TFrom, TTo> result)
            {
                if (predicate(value))
                    return Complete(result(value));

                return this;
            }

            public TTo Else(Func<TFrom, TTo> coerce)
            {
                return coerce(value);
            }

            public TTo ElseThrow()
            {
                throw new ApplicationException("Could not coerce {0} [{1}] to {2}".FormatWith((value?.GetType() ?? typeof(TFrom)).Name, value, typeof(TTo).Name));
            }

            public ICoercionContext<TFrom> Null(Func<TTo> result)
            {
                if (value == null)
                    return Complete(result());

                return this;
            }

            CoercionComplete<TFrom> Complete(TTo result)
            {
                return new CoercionComplete<TFrom>(result);
            }
        }
    }
}
