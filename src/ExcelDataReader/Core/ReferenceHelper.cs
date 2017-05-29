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
        public static void ParseReference(string value, out int column, out int row)
        {
            // INFO: Check for a simple Solution
            int index = ParseReference(value, out column);

            row = int.Parse(value.Substring(index), NumberStyles.None, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Logic for the Excel dimensions. Ex: A15
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="column">The column, 1-based.</param>
        /// <returns>The index of the row.</returns>
        public static int ParseReference(string value, out int column)
        {
            int index = 0;
            column = 0;

            const int offset = 'A' - 1;

            for (; index < value.Length; index++)
            {
                char c = value[index];
                if (char.IsDigit(c))
                    break;
                column *= 26;
                column += c - offset;
            }

            return index;
        }
    }
}
