namespace ExcelDataReader.Silverlight.Core
{
    using System;
    using System.Text;

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
            //return false;
            //CHANGED
            if (encoding == Encoding.Unicode)
                return true;
            else
                return false;
        }
#endif

        public static double Int64BitsToDouble(long value)
        {
            return BitConverter.ToDouble(BitConverter.GetBytes(value), 0);
        }

    }
}