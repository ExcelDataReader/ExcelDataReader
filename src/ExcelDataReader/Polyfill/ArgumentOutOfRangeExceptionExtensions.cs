#if !NET8_0_OR_GREATER
using System.Runtime.CompilerServices;

namespace System;

internal static class ArgumentOutOfRangeExceptionExtensions
{
    extension(ArgumentOutOfRangeException)
    {
        public static void ThrowIfLessThan<T>(T value, T other, [CallerArgumentExpression(nameof(value))] string? paramName = null)
            where T : IComparable<T>
        {
            if (value.CompareTo(other) < 0)
                throw new ArgumentOutOfRangeException(paramName, value, null);
        }
    }
}
#endif
