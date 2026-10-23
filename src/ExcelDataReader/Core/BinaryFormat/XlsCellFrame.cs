using System.Runtime.CompilerServices;

namespace ExcelDataReader.Core.BinaryFormat;

// A frame views either reader-borrowed bytes or an owned record payload.
// Borrowed bytes must be consumed before the next worksheet read.
internal readonly struct XlsCellFrame(byte[] bytes)
{
    public bool IsValid => bytes != null;

    public BIFFRECORDTYPE Id
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => (BIFFRECORDTYPE)BitConverter.ToUInt16(bytes, 0);
    }

    public int RecordSize
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => BitConverter.ToUInt16(bytes, 2);
    }

    public ushort RowIndex
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => BitConverter.ToUInt16(bytes, 4);
    }

    public ushort ColumnIndex
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => BitConverter.ToUInt16(bytes, 6);
    }

    public ushort XFormat
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => IsBiff2Cell ? (ushort)(bytes[8] & 0x3f) : BitConverter.ToUInt16(bytes, 8);
    }

    public ushort Format
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => IsBiff2Cell ? (ushort)(bytes[9] & 0x3f) : (ushort)0;
    }

    public bool IsBiff2Cell
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => Id is BIFFRECORDTYPE.NUMBER_OLD or BIFFRECORDTYPE.INTEGER_OLD or BIFFRECORDTYPE.LABEL_V2 or BIFFRECORDTYPE.BLANK_OLD or BIFFRECORDTYPE.BOOLERR_OLD;
    }

    public double NumberValue
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => BitConverter.ToDouble(bytes, Id == BIFFRECORDTYPE.NUMBER_OLD ? 11 : 10);
    }

    public uint SharedStringIndex
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => BitConverter.ToUInt32(bytes, 10);
    }

    public double RkValue
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => XlsBiffRKCell.NumFromRK(BitConverter.ToUInt32(bytes, 10));
    }

    public ushort LastColumnIndex
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => BitConverter.ToUInt16(bytes, RecordSize + 2);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ushort GetMultiCellStyle(ushort column)
    {
        int offset = 4 + 6 * (column - ColumnIndex);
        return offset > RecordSize - 2 ? (ushort)0 : BitConverter.ToUInt16(bytes, 4 + offset);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public double GetMultiCellValue(ushort column)
    {
        int offset = 6 + 6 * (column - ColumnIndex);
        return offset > RecordSize ? 0 : XlsBiffRKCell.NumFromRK(BitConverter.ToUInt32(bytes, 4 + offset));
    }
}
