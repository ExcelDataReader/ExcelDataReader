namespace ExcelDataReader.Silverlight.Core.BinaryFormat
{
	/// <summary>
	/// Represents BIFF EOF resord
	/// </summary>
	internal class XlsBiffEOF : XlsBiffRecord
	{
		internal XlsBiffEOF(byte[] bytes, uint offset)
			: base(bytes, offset)
		{
		}

		internal XlsBiffEOF(byte[] bytes)
			: this(bytes, 0)
		{
		}
	}
}
