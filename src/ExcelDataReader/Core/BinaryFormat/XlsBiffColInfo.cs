using System;

namespace ExcelDataReader.Core.BinaryFormat
{
    internal class XlsBiffColInfo : XlsBiffRecord
    {
        public XlsBiffColInfo(byte[] bytes, uint offset)
            : base(bytes, offset)
        {
            var colFirst = ReadUInt16(0x0);
            var colLast = ReadUInt16(0x2);
            var colDx = ReadUInt16(0x4);
            var flags = (ColInfoSettings)ReadUInt16(0x8);
            var userSet = (flags & ColInfoSettings.UserSet) != 0;
            var hidden = (flags & ColInfoSettings.Hidden) != 0;

            Value = new Col
            {
                CustomWidth = userSet,
                Max = colLast,
                Min = colFirst,
                Width = (double)colDx / 256,
                Hidden = hidden
            };
        }

        [Flags]
        private enum ColInfoSettings
        {
            Hidden = 0x01,
            UserSet = 0x10
        }

        public Col Value { get; }
    }
}
