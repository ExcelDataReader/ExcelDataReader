using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Sheet record in Workbook Globals
    /// </summary>
    internal class XlsBiffBoundSheet : XlsBiffRecord
    {
        internal XlsBiffBoundSheet(byte[] bytes, uint offset, bool isV8, Encoding encoding)
            : base(bytes, offset)
        {
            IsV8 = isV8;
            SheetNameEncoding = encoding;
        }

        public enum SheetType : byte
        {
            Worksheet = 0x0,
            MacroSheet = 0x1,
            Chart = 0x2,

            // ReSharper disable once InconsistentNaming
            VBModule = 0x6
        }

        public enum SheetVisibility : byte
        {
            Visible = 0x0,
            Hidden = 0x1,
            VeryHidden = 0x2
        }

        /// <summary>
        /// Gets the worksheet data start offset.
        /// </summary>
        public uint StartOffset => ReadUInt32(0x0);

        /// <summary>
        /// Gets the worksheet type.
        /// </summary>
        public SheetType Type => (SheetType)ReadByte(0x5);

        /// <summary>
        /// Gets the visibility of the worksheet.
        /// </summary>
        public SheetVisibility VisibleState => (SheetVisibility)ReadByte(0x4);

        /// <summary>
        /// Gets the name of the worksheet.
        /// </summary>
        public string SheetName
        {
            get
            {
                ushort len = ReadByte(0x6);

                const int start = 0x8;
                if (!IsV8)
                    return SheetNameEncoding.GetString(Bytes, RecordContentOffset + start, Helpers.IsSingleByteEncoding(SheetNameEncoding) ? len : len * 2);

                if (ReadByte(0x7) == 0)
                {
                    byte[] bytes = new byte[len * 2];
                    for (int i = 0; i < len; i++)
                    {
                        bytes[i * 2] = Bytes[RecordContentOffset + start + i];
                    }

                    return Encoding.Unicode.GetString(bytes, 0, len * 2);
                }

                return Encoding.Unicode.GetString(Bytes, RecordContentOffset + start, len * 2);
            }
        }

        /// <summary>
        /// Gets a value indicating whether this is a BIFF8 file or not.
        /// </summary>
        public bool IsV8 { get; }

        public Encoding SheetNameEncoding { get; }
    }
}