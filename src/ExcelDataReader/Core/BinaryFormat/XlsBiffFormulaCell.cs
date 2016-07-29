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

		private Encoding m_UseEncoding = Encoding.Unicode;

		internal XlsBiffFormulaCell(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
		{
		}

        ///// <summary>
        ///// Encoding used to deal with strings
        ///// </summary>
        //public Encoding UseEncoding
        //{
        //    get { return m_UseEncoding; }
        //    set { m_UseEncoding = value; }
        //}

		/// <summary>
		/// Formula flags
		/// </summary>
		public FormulaFlags Flags
		{
			get { return (FormulaFlags)(base.ReadUInt16(0xE)); }
		}

		/// <summary>
		/// Length of formula string
		/// </summary>
		public byte FormulaLength
		{
			get { return base.ReadByte(0xF); }
		}

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
                            XlsBiffRecord rec = GetRecord(m_bytes, (uint)(Offset + Size), reader);
                            XlsBiffFormulaString str;
                            if (rec.ID == BIFFRECORDTYPE.SHAREDFMLA)
								str = GetRecord(m_bytes, (uint)(Offset + Size + rec.Size), reader) as XlsBiffFormulaString;
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