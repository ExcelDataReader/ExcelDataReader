using System;
using System.Collections.Generic;
using System.Text;
using ExcelDataReader.Core.BinaryFormat;

namespace Excel.Core.BinaryFormat
{
    internal class XlsBiffBoolErr : XlsBiffBlankCell
    {
        internal XlsBiffBoolErr(byte[] bytes)
			: this(bytes, 0)
		{

		}

        internal XlsBiffBoolErr(byte[] bytes, uint offset)
			: base(bytes, offset)
		{

		}

        public bool BoolValue
        {
            get { return this.ReadByte(0x6) == 1; }
        }
    }
}
