using System.Runtime.CompilerServices;

namespace ExcelDataReader.Core.BinaryFormat;

// A frame views one record in the transport's reusable buffer and is valid
// only until the next read.
internal readonly struct XlsBiffFrame(byte[] bytes)
{
    public bool IsValid => bytes != null;

    public byte[] Bytes => bytes;

    public BIFFRECORDTYPE Id
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => (BIFFRECORDTYPE)BitConverter.ToUInt16(bytes, 0);
    }

    public bool IsCell => Id switch
    {
        BIFFRECORDTYPE.NUMBER or BIFFRECORDTYPE.NUMBER_OLD or
        BIFFRECORDTYPE.LABELSST or BIFFRECORDTYPE.RK or BIFFRECORDTYPE.MULRK or
        BIFFRECORDTYPE.BOOLERR or BIFFRECORDTYPE.BOOLERR_OLD or
        BIFFRECORDTYPE.INTEGER or BIFFRECORDTYPE.INTEGER_OLD or
        BIFFRECORDTYPE.LABEL or BIFFRECORDTYPE.LABEL_V2 or BIFFRECORDTYPE.RSTRING or
        BIFFRECORDTYPE.FORMULA or BIFFRECORDTYPE.FORMULA_V3 or BIFFRECORDTYPE.FORMULA_V4 or
        BIFFRECORDTYPE.BLANK or BIFFRECORDTYPE.BLANK_OLD or BIFFRECORDTYPE.MULBLANK => true,
        _ => false,
    };

    public bool IsEmptyCell => Id is BIFFRECORDTYPE.BLANK or BIFFRECORDTYPE.BLANK_OLD or BIFFRECORDTYPE.MULBLANK;

    public ushort RecordSize
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => BitConverter.ToUInt16(bytes, 2);
    }

    public ushort RowIndex
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => ReadUInt16At(4);
    }

    public ushort ColumnIndex
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => ReadUInt16At(6);
    }

    public ushort XFormat
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => IsBiff2Cell ? (ushort)(ReadByteAt(8) & 0x3f) : ReadUInt16At(8);
    }

    public ushort Format
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => IsBiff2Cell ? (ushort)(ReadByteAt(9) & 0x3f) : (ushort)0;
    }

    public bool IsBiff2Cell
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => Id is BIFFRECORDTYPE.NUMBER_OLD or BIFFRECORDTYPE.INTEGER_OLD or BIFFRECORDTYPE.LABEL_V2 or BIFFRECORDTYPE.BLANK_OLD or BIFFRECORDTYPE.BOOLERR_OLD;
    }

    public double NumberValue
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => ReadDoubleAt(Id == BIFFRECORDTYPE.NUMBER_OLD ? 11 : 10);
    }

    public uint SharedStringIndex
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => ReadUInt32At(10);
    }

    public double RkValue
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => DecodeRk(ReadUInt32At(10));
    }

    public ushort LastColumnIndex
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => ReadUInt16At(RecordSize + 2);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public byte ReadByte(int offset) => ReadByteAt(4 + offset);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ushort ReadUInt16(int offset) => ReadUInt16At(4 + offset);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public uint ReadUInt32(int offset) => ReadUInt32At(4 + offset);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public double ReadDouble(int offset) => ReadDoubleAt(4 + offset);

    public short ReadInt16(int offset) => unchecked((short)ReadUInt16(offset));

    public void ReadRowHeight(out bool useDefaultHeight, out int height)
    {
        var heightBits = ReadUInt16(6);
        useDefaultHeight = (heightBits & 0x8000) != 0;
        height = heightBits & 0x7FFFF;
        if (Id != BIFFRECORDTYPE.ROW_V2 && (ReadUInt16(12) & 0x20) != 0)
            height = 0;
    }

    public byte ReadByteAt(int offset)
    {
        EnsureRange(offset, sizeof(byte));
        return bytes[offset];
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201:Do not raise reserved exception types", Justification = "Preserve the legacy string decoder's exception for a truncated string header.")]
    public byte ReadStringByteAt(int offset)
    {
        if ((uint)offset >= RecordSize + 4)
            throw new IndexOutOfRangeException();

        return bytes[offset];
    }

    public ushort ReadUInt16At(int offset)
    {
        EnsureRange(offset, sizeof(ushort));
        return BitConverter.ToUInt16(bytes, offset);
    }

    public uint ReadUInt32At(int offset)
    {
        EnsureRange(offset, sizeof(uint));
        return BitConverter.ToUInt32(bytes, offset);
    }

    public double ReadDoubleAt(int offset)
    {
        EnsureRange(offset, sizeof(double));
        return BitConverter.ToDouble(bytes, offset);
    }

    public void EnsureRange(int offset, int count)
    {
        if (offset < 0 || count < 0 || offset > RecordSize + 4 - count)
            throw new ArgumentException(Errors.ErrorBiffRecordSize);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ushort GetMultiCellStyle(ushort column)
    {
        int offset = 4 + 6 * (column - ColumnIndex);
        return offset > RecordSize - 2 ? (ushort)0 : ReadUInt16At(4 + offset);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public double GetMultiCellValue(ushort column)
    {
        int offset = 6 + 6 * (column - ColumnIndex);
        return offset > RecordSize ? 0 : DecodeRk(ReadUInt32At(4 + offset));
    }

    private static double DecodeRk(uint rk)
    {
        double number = (rk & 2) != 0
            ? (int)rk >> 2
            : BitConverter.Int64BitsToDouble((long)(rk & 0xfffffffc) << 32);
        return (rk & 1) != 0 ? number / 100 : number;
    }
}