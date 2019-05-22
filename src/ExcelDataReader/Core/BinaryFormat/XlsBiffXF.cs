using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    [Flags]
    internal enum XfUsedAttributes : byte
    {
        NumberFormat = 0x01,
        Font = 0x02,
        TextStyle = 0x04,
        BorderLines = 0x08,
        BackgroundAreaStyle = 0x10,
        CellProtection = 0x20,
    }

    internal class XlsBiffXF : XlsBiffRecord
    {
        internal XlsBiffXF(byte[] bytes, uint offset, int biffVersion)
            : base(bytes, offset)
        {
            switch (Id)
            {
                case BIFFRECORDTYPE.XF_V2:
                    Format = ReadByte(2) & 0x3F;
                    HorizontalAlignment = (HorizontalAlignment)(ReadByte(3) & 0x07);
                    break;
                case BIFFRECORDTYPE.XF_V3:
                    Format = ReadByte(1);
                    UsedAttributes = (XfUsedAttributes)(ReadByte(3) >> 2);
                    Parent = ReadUInt16(4) >> 4;
                    HorizontalAlignment = (HorizontalAlignment)(ReadByte(4) & 0x07);
                    break;
                case BIFFRECORDTYPE.XF_V4:
                    Format = ReadByte(1);
                    Parent = ReadUInt16(2) >> 4;
                    UsedAttributes = (XfUsedAttributes)(ReadByte(5) >> 2);
                    HorizontalAlignment = (HorizontalAlignment)(ReadByte(4) & 0x07);
                    break;
                default:
                    Format = ReadUInt16(2);
                    Parent = ReadUInt16(4) >> 4;
                    HorizontalAlignment = (HorizontalAlignment)(ReadByte(6) & 0x07);
                    if (biffVersion < 8)
                    {
                        UsedAttributes = (XfUsedAttributes)(ReadByte(7) >> 2);
                    }
                    else if (biffVersion == 8)
                    {
                        IndentLevel = ReadByte(8) & 0x0F;
                        UsedAttributes = (XfUsedAttributes)(ReadByte(9) >> 2);
                    }

                    break;
            }
        }

        public XfUsedAttributes UsedAttributes { get; }

        public int Format { get; }

        public int Parent { get; }

        public int IndentLevel { get; }

        public HorizontalAlignment HorizontalAlignment { get; }
    }
}
