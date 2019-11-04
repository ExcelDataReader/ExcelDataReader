using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader.Core.OpenXmlFormat.Records;

#nullable enable

namespace ExcelDataReader.Core.OpenXmlFormat.BinaryFormat
{
    internal sealed class BiffWorkbookReader : BiffReader
    {
        private const int WorkbookPr = 0x99;
        private const int Sheet = 0x9C;

        public BiffWorkbookReader(Stream stream)
            : base(stream)
        {
        }

        private enum SheetVisibility : byte
        {
            Visible = 0x0,
            Hidden = 0x1,
            VeryHidden = 0x2
        }

        protected override Record ReadOverride(byte[] buffer, uint recordId, uint recordLength)
        {
            switch (recordId)
            {
                case WorkbookPr:
                    return new WorkbookPrRecord((buffer[0] & 0x01) == 1);
                case Sheet: // BrtBundleSh
                    var state = (SheetVisibility)GetDWord(buffer, 0) switch
                    {
                        SheetVisibility.Hidden => "hidden",
                        SheetVisibility.VeryHidden => "veryhidden",
                        _ => "visible"
                    };

                    uint id = GetDWord(buffer, 4);

                    uint offset = 8;
                    string? rid = GetNullableString(buffer, ref offset);

                    // Must be between 1 and 31 characters
                    uint nameLength = GetDWord(buffer, offset);
                    string name = GetString(buffer, offset + 4, nameLength);

                    return new SheetRecord(name, id, rid, state);
                default:
                    return Record.Default;
            }
        }
    }
}
