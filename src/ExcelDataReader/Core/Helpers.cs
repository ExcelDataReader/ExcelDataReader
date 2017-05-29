using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDataReader.Misc;

namespace ExcelDataReader.Core
{
    /// <summary>
    /// Helpers class
    /// </summary>
    internal static class Helpers
    {
        private static readonly Regex EscapeRegex = new Regex("_x([0-9A-F]{4,4})_");

        /// <summary>
        /// Determines whether [is single byte] [the specified encoding].
        /// </summary>
        /// <param name="encoding">The encoding.</param>
        /// <returns>
        ///     <c>true</c> if [is single byte] [the specified encoding]; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsSingleByteEncoding(Encoding encoding)
        {
            return encoding.GetByteCount(new[] { 'a' }) == 1;
        }

        public static double Int64BitsToDouble(long value)
        {
            return BitConverter.ToDouble(BitConverter.GetBytes(value), 0);
        }

        public static string ConvertEscapeChars(string input)
        {
            return EscapeRegex.Replace(input, m => ((char)uint.Parse(m.Groups[1].Value, NumberStyles.HexNumber)).ToString());
        }

        public static object ConvertFromOATime(double value)
        {
            if (value >= 0.0 && value < 60.0)
            {
                value++;
            }

            /*
            if (date1904)
            {
                Value += 1462.0;
            }
            */
            
            return DateTimeHelper.FromOADate(value);
        }
    }
}