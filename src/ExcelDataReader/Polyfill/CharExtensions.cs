#if !NET8_0_OR_GREATER
namespace System;

internal static class CharExtensions
{
    extension(char)
    {
        public static bool IsAsciiDigit(char value) => value is >= '0' and <= '9';

        public static bool IsAsciiLetter(char value) => value is (>= 'A' and <= 'Z') or (>= 'a' and <= 'z');
    }
}
#endif
