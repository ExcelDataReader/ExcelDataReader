﻿namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// [MS-XLS] 2.4.168 MergeCells
///  If the count of the merged cells in the document is greater than 1026, the file will contain multiple adjacent MergeCells records.
/// </summary>
internal sealed class XlsBiffMergeCells : XlsBiffRecord
{
    public XlsBiffMergeCells(byte[] bytes)
        : base(bytes)
    {
        var count = ReadUInt16(0);

        MergeCells = [];
        for (int i = 0; i < count; i++)
        {
            var fromRow = ReadInt16(2 + i * 8 + 0);
            var toRow = ReadInt16(2 + i * 8 + 2);
            var fromCol = ReadInt16(2 + i * 8 + 4);
            var toCol = ReadInt16(2 + i * 8 + 6);

            CellRange mergeCell = new(fromCol, fromRow, toCol, toRow);
            MergeCells.Add(mergeCell);
        }
    }

    public List<CellRange> MergeCells { get; }
}
