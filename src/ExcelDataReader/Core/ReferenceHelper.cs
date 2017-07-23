using System.Globalization;

namespace ExcelDataReader.Core
{
    internal static class ReferenceHelper
    {
        /// <summary>
        /// Logic for the Excel dimensions. Ex: A15
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

                    if (char.IsDigit(c))
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

            if (!int.TryParse(value.Substring(position), NumberStyles.None, CultureInfo.InvariantCulture, out row))
            {
                return false;
            }

            return true;
        }
    }
}
