#if !(NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER)
namespace System;

internal static class MathExtensions
{
    extension(Math)
    {
        public static long Clamp(long value, long min, long max)
        {
            if (min > max)
                throw new ArgumentException("Minimum must not be greater than maximum.");
            return Math.Min(max, Math.Max(min, value));
        }
    }
}
#endif
