using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Log
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
