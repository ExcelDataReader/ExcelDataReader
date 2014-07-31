namespace ExcelDataReader.Silverlight.Core.BinaryFormat
{

	/// <summary>
	/// Represents InterfaceHdr record in Wokrbook Globals
	/// </summary>
	internal class XlsBiffInterfaceHdr : XlsBiffRecord
	{
		internal XlsBiffInterfaceHdr(byte[] bytes, uint offset)
			: base(bytes, offset)
		{
		}

		internal XlsBiffInterfaceHdr(byte[] bytes)
			: this(bytes, 0)
		{
		}

		/// <summary>
		/// Returns CodePage for Interface Header
		/// </summary>
		public ushort CodePage
		{
			get { return base.ReadUInt16(0x0); }
		}
	}
}
