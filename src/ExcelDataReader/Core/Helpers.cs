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
        /// Determines whether the encoding is single byte or not.
        /// </summary>
        /// <param name="encoding">The encoding.</param>
        /// <returns>
        ///     <see langword="true"/> if the specified encoding is single byte; otherwise, <see langword="false"/>.
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

        /// <summary>
        /// Convert a double from Excel to an OA DateTime double. 
        /// The returned value is normalized to the '1900' date mode and adjusted for the 1900 leap year bug.
        /// </summary>
        public static double AdjustOADateTime(double value, bool date1904)
        {
            if (!date1904)
            {
                // Workaround for 1900 leap year bug in Excel
                if (value >= 0.0 && value < 60.0)
                {
                    return value + 1;
                }
            }
            else
            {
                return value + 1462.0;
            }

            return value;
        }

        public static bool IsValidOADateTime(double value)
        {
            return value > DateTimeHelper.OADateMinAsDouble && value < DateTimeHelper.OADateMaxAsDouble;
        }

        public static object ConvertFromOATime(double value, bool date1904)
        {
            var dateValue = AdjustOADateTime(value, date1904);
            if (IsValidOADateTime(dateValue))
                return DateTimeHelper.FromOADate(dateValue);
            return value;
        }

        public static object ConvertFromOATime(int value, bool date1904)
        {
            var dateValue = AdjustOADateTime(value, date1904);
            if (IsValidOADateTime(dateValue))
                return DateTimeHelper.FromOADate(dateValue);
            return value;
        }
    }
}