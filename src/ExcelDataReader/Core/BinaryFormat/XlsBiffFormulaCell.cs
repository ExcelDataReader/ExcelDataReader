using System;
using System.Text;

using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
	/// <summary>
	/// Represents a cell containing formula
	/// </summary>
	internal class XlsBiffFormulaCell : XlsBiffNumberCell
	{
		#region FormulaFlags enum

		[Flags]
		public enum FormulaFlags : ushort
		{
			AlwaysCalc = 0x0001,
			CalcOnLoad = 0x0002,
			SharedFormulaGroup = 0x0008
		}

		#endregion

		internal XlsBiffFormulaCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
		{
		}

		/// <summary>
		/// Formula flags
		/// </summary>
		public FormulaFlags Flags => (FormulaFlags)ReadUInt16(0xE);

	    /// <summary>
		/// Length of formula string
		/// </summary>
		public byte FormulaLength => ReadByte(0xF);

	    /// <summary>
		/// Returns type-dependent value of formula
		/// </summary>
		public new object Value
		{
			get
			{
				long val = base.ReadInt64(0x6);
				if (((ulong)val & 0xFFFF000000000000) == 0xFFFF000000000000)
				{
					byte type = (byte)(val & 0xFF);
					byte code = (byte)((val >> 16) & 0xFF);
					switch (type)
					{
						case 0: // String

                            //////////////fix
                            XlsBiffRecord rec = GetRecord(Bytes, (uint)(Offset + Size), Reader);
                            XlsBiffFormulaString str;
                            if (rec.ID == BIFFRECORDTYPE.SHAREDFMLA)
								str = GetRecord(Bytes, (uint)(Offset + Size + rec.Size), Reader) as XlsBiffFormulaString;
                            else
                                str = rec as XlsBiffFormulaString;
                            //////////////fix

                            if (str == null)
                                return string.Empty;
                            else
                            {
                                //str.UseEncoding = m_UseEncoding;
                                return str.Value;
                            }
						case 1: // Boolean

							return (code != 0);
						case 2: // Error

							return (FORMULAERROR)code;
						default:
							return null;
					}
				}
				else
					return Helpers.Int64BitsToDouble(val);
			}
		}

		public string Formula
		{
			get
			{
				byte[] bts = base.ReadArray(0x10, FormulaLength);
                return Encoding.Unicode.GetString(bts, 0, bts.Length);
			}
		}
	}
}