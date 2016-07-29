
using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
	/// <summary>
	/// Represents record with the only two-bytes value
	/// </summary>
	internal class XlsBiffSimpleValueRecord : XlsBiffRecord
	{
		internal XlsBiffSimpleValueRecord(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
		{
		}

		/// <summary>
		/// Returns value
		/// </summary>
		public ushort Value
		{
			get { return ReadUInt16(0x0); }
		}
	}
}
