namespace ExcelDataReader.Silverlight.Core.BinaryFormat
{
	/// <summary>
	/// Represents additional space for very large records
	/// </summary>
	internal class XlsBiffContinue : XlsBiffRecord
	{
		internal XlsBiffContinue(byte[] bytes, uint offset)
			: base(bytes, offset)
		{
		}

		internal XlsBiffContinue(byte[] bytes)
			: this(bytes, 0)
		{
		}
	}
}
