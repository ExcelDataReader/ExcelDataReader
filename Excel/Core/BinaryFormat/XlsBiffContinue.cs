namespace Excel.Core.BinaryFormat
{
	/// <summary>
	/// Represents additional space for very large records
	/// </summary>
	internal class XlsBiffContinue : XlsBiffRecord
	{
		internal XlsBiffContinue(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
		{
		}

	}
}
