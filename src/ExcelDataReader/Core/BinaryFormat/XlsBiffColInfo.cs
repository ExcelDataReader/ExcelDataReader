using System;

namespace ExcelDataReader.Core.BinaryFormat
{
    internal class XlsBiffColInfo : XlsBiffRecord
    {
        public XlsBiffColInfo(byte[] bytes)
            : base(bytes)
        {
            var colFirst = ReadUInt16(0x0);
            var colLast = ReadUInt16(0x2);
            var colDx = ReadUInt16(0x4);
            var flags = (ColInfoSettings)ReadUInt16(0x8);
            var userSet = (flags & ColInfoSettings.UserSet) != 0;
            var hidden = (flags & ColInfoSettings.Hidden) != 0;

            Value = new Column(colFirst, colLast, hidden, userSet ? (double?)colDx / 256.0 : null);
        }

        [Flags]
        private enum ColInfoSettings
        {
            Hidden = 0b01,
            UserSet = 0b10,
        }

        public Column Value { get; }
    }
}
