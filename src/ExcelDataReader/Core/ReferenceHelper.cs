using System;
using System.Globalization;

namespace ExcelDataReader.Core
{
    internal static class ReferenceHelper
    {
        /// <summary>
        /// Logic for the Excel dimensions. Ex: A15.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="column">The column, 1-based.</param>
        /// <param name="row">The row, 1-based.</param>
        public static bool ParseReference(string value, out int column, out int row)
        {
            column = 0;
            var position = 0;
            const int offset = 'A' - 1;

            if (value != null)
            {
                while (position < value.Length)
                {
                    var c = value[position];
                    if (c >= 'A' && c <= 'Z')
                    {
                        position++;
                        column *= 26;
                        column += c - offset;
                        continue;
                    }

                    if (IsDigit(c))
                        break;

                    position = 0;
                    break;
                }
            }

            if (position == 0)
            {
                column = 0;
                row = 0;
                return false;
            }

            if (!TryParseDecInt(value, position, out row))
            {
                return false;
            }

            return true;
        }

        private static bool IsDigit(int ch) => ((uint)ch - '0') <= 9;

        private static bool TryParseDecInt(string s, int startIndex, out int result)
        {
            if (startIndex >= s.Length)
                goto Fail;

            long r = 0;
            for (int i = startIndex; i < s.Length; ++i)
            {
                int num = s[i];
                if (!IsDigit(num))
                    goto Fail;
                r = r * 10 + (num - '0');
                if (r > int.MaxValue)
                    goto Fail;
            }

            result = (int)r;
            return true;

        Fail:
            result = 0;
            return false;
        }
    }
}
