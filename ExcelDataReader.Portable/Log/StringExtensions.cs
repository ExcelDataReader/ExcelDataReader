namespace ExcelDataReader.Portable.Log
{
	public static class StringExtensions
	{
		/// <summary>
		/// Formats string with the formatting passed in. This is a shortcut to string.Format().
		/// </summary>
		/// <param name="input">The input.</param>
		/// <param name="formatting">The formatting.</param>
		/// <returns>A formatted string.</returns>
		public static string FormatWith(this string input, params object[] formatting)
		{
			return string.Format(input, formatting);
		}

	}
}
