using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat
{
    internal sealed class BiffWorksheetReader : BiffReader
    {
        private const uint Row = 0x00; 
        private const uint Blank = 0x01;
        private const uint Number = 0x02;
        private const uint BoolError = 0x03; 
        private const uint Bool = 0x04; 
        private const uint Float = 0x05;
        private const uint String = 0x06;
        private const uint SharedString = 0x07;
        private const uint FormulaString = 0x08;
        private const uint FormulaNumber = 0x09;
        private const uint FormulaBool = 0x0a;
        private const uint FormulaError = 0x0b;

        // private const uint WorksheetBegin = 0x81;
        // private const uint WorksheetEnd = 0x82;
        private const uint SheetDataBegin = 0x91;
        private const uint SheetDataEnd = 0x92;
        private const uint SheetPr = 0x93;
        private const uint SheetFormatPr = 0x1E5;

        // private const uint ColumnsBegin = 0x186;
        private const uint Column = 0x3C; // column info

        // private const uint ColumnsEnd = 0x187;
        private const uint HeaderFooter = 0x1DF;

        // private const uint MergeCellsBegin = 177;
        // private const uint MergeCellsEnd = 178;
        private const uint MergeCell = 176;

        public BiffWorksheetReader(Stream stream) 
            : base(stream)
        {
        }

        protected override Record ReadOverride(byte[] buffer, uint recordId, uint recordLength)
        {
            switch (recordId) 
            {
                case SheetDataBegin:
                    return new SheetDataBeginRecord();
                case SheetDataEnd:
                    return new SheetDataEndRecord();
                case SheetPr: // BrtWsProp
                    {
                        // Must be between 0 and 31 characters
                        uint length = GetDWord(buffer, 19);

                        // To behave the same as when reading an xml based file. 
                        // GetAttribute returns null both if the attribute is missing
                        // or if it is empty.
                        string codeName = length == 0 ? null : GetString(buffer, 19 + 4, length);
                        return new SheetPrRecord(codeName);
                    }

                case SheetFormatPr: // BrtWsFmtInfo 
                    {
                        // TODO Default column width
                        var unsynced = (buffer[8] & 0b1000) != 0;
                        uint? defaultHeight = null;
                        if (unsynced)
                            defaultHeight = GetWord(buffer, 6);
                        return new SheetFormatPrRecord(defaultHeight);
                    }

                case Column: // BrtColInfo 
                    {
                        int minimum = GetInt32(buffer, 0);
                        int maximum = GetInt32(buffer, 4);
                        byte flags = buffer[16];
                        bool hidden = (flags & 0b1) != 0;
                        bool unsynced = (flags & 0b10) != 0;

                        double? width = null;
                        if (unsynced)
                            width = GetDWord(buffer, 8) / 256.0;
                        return new ColumnRecord(new Column(minimum, maximum, hidden, width));
                    }

                case HeaderFooter: // BrtBeginHeaderFooter 
                    {
                        var flags = buffer[0];
                        bool differentOddEven = (flags & 1) != 0;
                        bool differentFirst = (flags & 0b10) != 0;
                        uint offset = 2;
                        var header = GetNullableString(buffer, ref offset);
                        var footer = GetNullableString(buffer, ref offset);
                        var headerEven = GetNullableString(buffer, ref offset);
                        var footerEven = GetNullableString(buffer, ref offset);
                        var headerFirst = GetNullableString(buffer, ref offset);
                        var footerFirst = GetNullableString(buffer, ref offset);
                        return new HeaderFooterRecord(new HeaderFooter(differentFirst, differentOddEven) 
                        {
                            FirstHeader = headerFirst,
                            FirstFooter = footerFirst,
                            OddHeader = header,
                            OddFooter = footer,
                            EvenHeader = headerEven,
                            EvenFooter = footerEven,
                        });
                    }

                case MergeCell:
                    int fromRow = GetInt32(buffer, 0);
                    int toRow = GetInt32(buffer, 4);
                    int fromColumn = GetInt32(buffer, 8);
                    int toColumn = GetInt32(buffer, 12);
                    return new MergeCellRecord(new CellRange(fromColumn, fromRow, toColumn, toRow));
                case Row: // BrtRowHdr 
                    {
                        int rowIndex = GetInt32(buffer, 0);
                        byte flags = buffer[11];
                        bool hidden = (flags & 0b10000) != 0;
                        bool unsynced = (flags & 0b100000) != 0;

                        double? height = null;
                        if (unsynced)
                            height = GetWord(buffer, 8) / 20.0; // Where does 20 come from?

                        // TODO: Default format ?
                        return new RowHeaderRecord(rowIndex, hidden, height);
                    }

                case Blank:
                    return ReadCell(null);
                case BoolError:
                case FormulaError:
                    return ReadCell(null, (CellError)buffer[8]);
                case Number:
                    return ReadCell(GetRkNumber(buffer, 8));
                case Bool:
                case FormulaBool:
                    return ReadCell(buffer[8] == 1);
                case FormulaNumber:
                case Float:
                    return ReadCell(GetDouble(buffer, 8));
                case String:
                case FormulaString:
                    {
                        // Must be less than 32768 characters
                        var length = GetDWord(buffer, 8);
                        return ReadCell(GetString(buffer, 8 + 4, length));
                    }

                case SharedString:
                    return ReadCell((int)GetDWord(buffer, 8));
                default:
                    return Record.Default;
            }

            CellRecord ReadCell(object value, CellError? errorValue = null) 
            {
                int column = (int)GetDWord(buffer, 0);
                uint xfIndex = GetDWord(buffer, 4) & 0xffffff;

                return new CellRecord(column, (int)xfIndex, value, errorValue);
            }
        }
    }
}
