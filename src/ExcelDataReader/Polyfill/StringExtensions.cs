#if !(NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER)
namespace System;

internal static class StringExtensions
{
    public static bool StartsWith(this string value, char start) =>
        value.Length > 0 && value[0] == start;
}
#endif
