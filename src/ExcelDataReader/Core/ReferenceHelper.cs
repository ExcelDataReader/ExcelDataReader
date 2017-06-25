using System;
using System.Globalization;

namespace ExcelDataReader.Core
{
    public static class ReferenceHelper
    {
        /// <summary>
        /// Logic for the Excel dimensions. Ex: A15
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="column">The column, 1-based.</param>
        /// <param name="row">The row, 1-based.</param>
        public static bool ParseReference(string value, out int column, out int row)
        {
            var position = 0;
            if (value != null)
            {
                while (position < value.Length)
                {
                    var c = value[position];
                    if (c >= 'A' && c <= 'Z')
                    {
                        position++;
                        continue;
                    }
                    else if (char.IsDigit(c))
                    {
                        break;
                    }
                    else
                    {
                        position = 0;
                        break;
                    }
                }
            }

            if (position == 0)
            {
                column = 0;
                row = 0;
                return false;
            }

            column = ParseColumn(value.Substring(0, position));

            if (!int.TryParse(value.Substring(position), NumberStyles.None, CultureInfo.InvariantCulture, out row))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Logic for the Excel column. Ex: A, XY
        /// </summary>
        /// <param name="value">The column string value.</param>
        /// <returns>The column, 1-based.</returns>
        private static int ParseColumn(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                throw new FormatException("Empty column reference");
            }

            int index = 0;
            int column = 0;

            const int offset = 'A' - 1;

            for (; index < value.Length; index++)
            {
                char c = value[index];
                if (c < 'A' || c > 'Z')
                {
                    throw new FormatException("Invalid character in column reference " + value);
                }

                column *= 26;
                column += c - offset;
            }

            return column;
        }
    }
}
