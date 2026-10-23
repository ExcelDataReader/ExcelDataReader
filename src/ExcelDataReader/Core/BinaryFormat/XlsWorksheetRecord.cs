using System.Runtime.CompilerServices;

namespace ExcelDataReader.Core.BinaryFormat;

// Cell frames borrow the reusable buffer unless OwnedRecord supplies their payload.
// A default frame marks an owned metadata record. Borrowed frames last until the next read.
internal readonly struct XlsWorksheetRecord(XlsCellFrame cell, XlsBiffRecord? ownedRecord)
{
    public XlsCellFrame Cell { get; } = cell;

    public XlsBiffRecord? OwnedRecord { get; } = ownedRecord;

    public bool IsCell => Cell.IsValid;

    public static XlsWorksheetRecord ForCell(XlsCellFrame cell, XlsBiffRecord? ownedRecord) =>
        new(cell, ownedRecord);

    public static XlsWorksheetRecord ForRecord(XlsBiffRecord record) =>
        new(default, record);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
#pragma warning disable CA1822 // Instance-bound only where buffers are pooled
    public void Release()
#pragma warning restore CA1822
    {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
        OwnedRecord?.Return();
#endif
    }
}
