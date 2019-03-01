using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader.Core.NumberFormat;
using ExcelDataReader.Core.OfficeCrypto;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Globals section of workbook
    /// </summary>
    internal class XlsWorkbook : CommonWorkbook, IWorkbook<XlsWorksheet>
    {
        internal XlsWorkbook(Stream stream, string password, Encoding fallbackEncoding)
        {
            Stream = stream;

            using (var biffStream = new XlsBiffStream(stream, 0, 0, password))
            {
                if (biffStream.BiffVersion == 0)
                    throw new ExcelReaderException(Errors.ErrorWorkbookGlobalsInvalidData);

                BiffVersion = biffStream.BiffVersion;
                SecretKey = biffStream.SecretKey;
                Encryption = biffStream.Encryption;
                Encoding = biffStream.BiffVersion == 8 ? Encoding.Unicode : fallbackEncoding;

                if (biffStream.BiffType == BIFFTYPE.WorkbookGlobals)
                {
                    ReadWorkbookGlobals(biffStream);
                }
                else if (biffStream.BiffType == BIFFTYPE.Worksheet)
                {
                    // set up 'virtual' bound sheet pointing at this
                    Sheets.Add(new XlsBiffBoundSheet(0, XlsBiffBoundSheet.SheetType.Worksheet, XlsBiffBoundSheet.SheetVisibility.Visible, "Sheet"));
                }
                else
                {
                    throw new ExcelReaderException(Errors.ErrorWorkbookGlobalsInvalidData);
                }
            }
        }

        public Stream Stream { get; }

        public int BiffVersion { get; }

        public byte[] SecretKey { get; }

        public EncryptionInfo Encryption { get; }

        public Encoding Encoding { get; private set; }

        public XlsBiffInterfaceHdr InterfaceHdr { get; set; }

        public XlsBiffRecord Mms { get; set; }

        public XlsBiffRecord WriteAccess { get; set; }

        public XlsBiffSimpleValueRecord CodePage { get; set; }

        public XlsBiffRecord Dsf { get; set; }

        public XlsBiffRecord Country { get; set; }

        public XlsBiffSimpleValueRecord Backup { get; set; }

        public List<XlsBiffRecord> Fonts { get; } = new List<XlsBiffRecord>();

        public List<XlsBiffBoundSheet> Sheets { get; } = new List<XlsBiffBoundSheet>();

        /// <summary>
        /// Gets or sets the Shared String Table of workbook
        /// </summary>
        public XlsBiffSST SST { get; set; }

        public XlsBiffRecord ExtSST { get; set; }

        public bool IsDate1904 { get; private set; }

        public int ResultsCount => Sheets?.Count ?? -1;

        public static bool IsRawBiffStream(byte[] bytes)
        {
            if (bytes.Length < 8)
            {
                throw new ArgumentException("Needs at least 8 bytes to probe", nameof(bytes));
            }

            ushort rid = BitConverter.ToUInt16(bytes, 0);
            ushort size = BitConverter.ToUInt16(bytes, 2);
            ushort bofVersion = BitConverter.ToUInt16(bytes, 4);
            ushort type = BitConverter.ToUInt16(bytes, 6);

            switch (rid)
            {
                case 0x0009: // BIFF2
                    if (size != 4)
                        return false;
                    if (type != 0x10 && type != 0x20 && type != 0x40)
                        return false;
                    return true;
                case 0x0209: // BIFF3
                case 0x0409: // BIFF4
                    if (size != 6)
                        return false;
                    if (type != 0x10 && type != 0x20 && type != 0x40 && type != 0x0100)
                        return false;
                    /* removed this additional check to keep the probe at 8 bytes
                    ushort notUsed = BitConverter.ToUInt16(bytes, 8);
                    if (notUsed != 0x00)
                        return false;*/
                    return true;
                case 0x0809: // BIFF5/BIFF8
                    if (size < 4)
                        return false;
                    if (bofVersion != 0 && bofVersion != 0x0200 && bofVersion != 0x0300 && bofVersion != 0x0400 && bofVersion != 0x0500 && bofVersion != 0x600)
                        return false;
                    if (type != 0x5 && type != 0x6 && type != 0x10 && type != 0x20 && type != 0x40 && type != 0x0100)
                        return false;
                    /* removed this additional check to keep the probe at 8 bytes
                    ushort identifier = BitConverter.ToUInt16(bytes, 10);
                    if (identifier == 0)
                        return false;*/
                    return true;
            }

            return false;
        }

        public IEnumerable<XlsWorksheet> ReadWorksheets()
        {
            for (var i = 0; i < Sheets.Count; ++i)
            {
                yield return new XlsWorksheet(this, Sheets[i], Stream);
            }
        }

        private void ReadWorkbookGlobals(XlsBiffStream biffStream)
        {
            XlsBiffRecord rec;
            var biffFormats = new Dictionary<ushort, XlsBiffFormatString>();

            while ((rec = biffStream.Read()) != null && rec.Id != BIFFRECORDTYPE.EOF)
            {
                switch (rec.Id)
                {
                    case BIFFRECORDTYPE.INTERFACEHDR:
                        InterfaceHdr = (XlsBiffInterfaceHdr)rec;
                        break;
                    case BIFFRECORDTYPE.BOUNDSHEET:
                        XlsBiffBoundSheet sheet = (XlsBiffBoundSheet)rec;

                        if (sheet.Type != XlsBiffBoundSheet.SheetType.Worksheet)
                            break;

                        Sheets.Add(sheet);
                        break;
                    case BIFFRECORDTYPE.MMS:
                        Mms = rec;
                        break;
                    case BIFFRECORDTYPE.COUNTRY:
                        Country = rec;
                        break;
                    case BIFFRECORDTYPE.CODEPAGE:
                        // [MS-XLS 2.4.52 CodePage] An unsigned integer that specifies the workbook’s code page.The value MUST be one
                        // of the code page values specified in [CODEPG] or the special value 1200, which means that the
                        // workbook is Unicode.
                        CodePage = (XlsBiffSimpleValueRecord)rec;
                        Encoding = EncodingHelper.GetEncoding(CodePage.Value);
                        break;
                    case BIFFRECORDTYPE.FONT:
                    case BIFFRECORDTYPE.FONT_V34:
                        Fonts.Add(rec);
                        break;
                    case BIFFRECORDTYPE.FORMAT_V23:
                        {
                            var fmt = (XlsBiffFormatString)rec;
                            biffFormats.Add((ushort)biffFormats.Count, fmt);
                        }

                        break;
                    case BIFFRECORDTYPE.FORMAT:
                        {
                            var fmt = (XlsBiffFormatString)rec;
                            biffFormats.Add(fmt.Index, fmt);
                        }

                        break;
                    case BIFFRECORDTYPE.XF:
                    case BIFFRECORDTYPE.XF_V4:
                    case BIFFRECORDTYPE.XF_V3:
                    case BIFFRECORDTYPE.XF_V2:
                        AddExtendedFormat(GetExtendedFormatCount(), ((XlsBiffXF)rec).Format, true);
                        break;
                    case BIFFRECORDTYPE.SST:
                        SST = (XlsBiffSST)rec;
                        SST.ReadStrings(biffStream);
                        break;
                    case BIFFRECORDTYPE.CONTINUE:
                        break;
                    case BIFFRECORDTYPE.EXTSST:
                        ExtSST = rec;
                        break;
                    case BIFFRECORDTYPE.PASSWORD:
                        break;
                    case BIFFRECORDTYPE.PROTECT:
                    case BIFFRECORDTYPE.PROT4REVPASSWORD:
                        // IsProtected
                        break;
                    case BIFFRECORDTYPE.RECORD1904:
                        IsDate1904 = ((XlsBiffSimpleValueRecord)rec).Value == 1;
                        break;
                    default:
                        break;
                }
            }

            foreach (var biffFormat in biffFormats)
            {
                AddNumberFormat(biffFormat.Key, biffFormat.Value.GetValue(Encoding));
            }
        }
    }
}