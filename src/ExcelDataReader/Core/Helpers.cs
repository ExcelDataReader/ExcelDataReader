using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Text;
using ExcelDataReader.Misc;

namespace ExcelDataReader.Core
{
	/// <summary>
	/// Helpers class
	/// </summary>
	internal static class Helpers
	{
#if CF_DEBUG || CF_RELEASE

		/// <summary>
		/// Determines whether [is single byte] [the specified encoding].
		/// </summary>
		/// <param name="encoding">The encoding.</param>
		/// <returns>
		/// 	<c>true</c> if [is single byte] [the specified encoding]; otherwise, <c>false</c>.
		/// </returns>
		public static bool IsSingleByteEncoding(Encoding encoding)
		{
			return encoding.GetChars(new byte[] { 0xc2, 0xb5 }).Length == 1;
		}
#else

		/// <summary>
		/// Determines whether [is single byte] [the specified encoding].
		/// </summary>
		/// <param name="encoding">The encoding.</param>
		/// <returns>
		/// 	<c>true</c> if [is single byte] [the specified encoding]; otherwise, <c>false</c>.
		/// </returns>
		public static bool IsSingleByteEncoding(Encoding encoding)
		{
			return encoding.GetByteCount(new char[]{'a'}) == 1;
		}
#endif

		public static double Int64BitsToDouble(long value)
		{
			return BitConverter.ToDouble(BitConverter.GetBytes(value), 0);
		}

	    private static Regex re = new Regex("_x([0-9A-F]{4,4})_");

        public static string ConvertEscapeChars(string input)
        {
            return re.Replace(input, m => (((char)UInt32.Parse(m.Groups[1].Value, NumberStyles.HexNumber))).ToString());
        }

	    public static object ConvertFromOATime(double value)
	    {
	        if ((value >= 0.0) && (value < 60.0))
	        {
	            value++;
	        }
	        //if (date1904)
	        //{
	        //    Value += 1462.0;
	        //}
	        return DateTimeHelper.FromOADate(value);
	    }
    }
}