using System;

namespace ExcelDataReader.Portable.Exceptions
{
	public class HeaderException : Exception
	{
		public HeaderException()
		{
		}

		public HeaderException(string message)
			: base(message)
		{
		}

		public HeaderException(string message, Exception innerException)
			: base(message, innerException)
		{
		}
	}
}
