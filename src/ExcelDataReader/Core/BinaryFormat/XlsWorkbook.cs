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

        public List<XlsBiffFont> Fonts { get; } = new List<XlsBiffFont>();

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

        internal void AddXf(XlsBiffXF xf)
        {
            var extendedFormat = new ExtendedFormat()
            {
                FontIndex = xf.Font,
                NumberFormatIndex = xf.Format,
                Locked = xf.IsLocked,
                Hidden = xf.IsHidden,
                HorizontalAlignment = xf.HorizontalAlignment,
                IndentLevel = xf.IndentLevel,
                ParentCellStyleXf = xf.ParentCellStyleXf,
            };

            // The workbook holds two kinds of XF records: Cell XFs, and Cell Style XFs.
            // In the binary XLS format, both kinds of XF records are saved in a single list,
            // whereas the XLSX format has two separate lists - like the CommonWorkbook internals.
            // The Cell XFs hold indexes into the Cell Style XF list, so adding the XF in both lists 
            // here to keep the indexes the same.
            ExtendedFormats.Add(extendedFormat);
            CellStyleExtendedFormats.Add(extendedFormat);
        }

        private void ReadWorkbookGlobals(XlsBiffStream biffStream)
        {
            var formats = new Dictionary<int, XlsBiffFormatString>();

            XlsBiffRecord rec;
            while ((rec = biffStream.Read()) != null && !(rec is XlsBiffEof))
            {
                switch (rec)
                {
                    case XlsBiffInterfaceHdr hdr:
                        InterfaceHdr = hdr;
                        break;
                    case XlsBiffBoundSheet sheet:
                        if (sheet.Type != XlsBiffBoundSheet.SheetType.Worksheet)
                            break;
                        Sheets.Add(sheet);
                        break;
                    case XlsBiffSimpleValueRecord codePage when rec.Id == BIFFRECORDTYPE.CODEPAGE:
                        // [MS-XLS 2.4.52 CodePage] An unsigned integer that specifies the workbook’s code page.The value MUST be one
                        // of the code page values specified in [CODEPG] or the special value 1200, which means that the
                        // workbook is Unicode.
                        CodePage = codePage;
                        Encoding = EncodingHelper.GetEncoding(CodePage.Value);
                        break;
                    case XlsBiffSimpleValueRecord is1904 when rec.Id == BIFFRECORDTYPE.RECORD1904:
                        IsDate1904 = is1904.Value == 1;
                        break;
                    case XlsBiffFont font:
                        Fonts.Add(font);
                        break;
                    case XlsBiffFormatString format23 when rec.Id == BIFFRECORDTYPE.FORMAT_V23:
                        formats.Add((ushort)formats.Count, format23);
                        break;
                    case XlsBiffFormatString fmt when rec.Id == BIFFRECORDTYPE.FORMAT:
                        var index = fmt.Index;
                        if (!formats.ContainsKey(index))
                            formats.Add(index, fmt);
                        break;
                    case XlsBiffXF xf:
                        AddXf(xf);
                        break;
                    case XlsBiffSST sst:
                        SST = sst;
                        break;
                    case XlsBiffContinue sstContinue:
                        if (SST != null)
                        {
                            SST.ReadContinueStrings(sstContinue);
                        }

                        break;
                    case XlsBiffRecord _ when rec.Id == BIFFRECORDTYPE.MMS:
                        Mms = rec;
                        break;
                    case XlsBiffRecord _ when rec.Id == BIFFRECORDTYPE.COUNTRY:
                        Country = rec;
                        break;
                    case XlsBiffRecord _ when rec.Id == BIFFRECORDTYPE.EXTSST:
                        ExtSST = rec;
                        break;

                    // case BIFFRECORDTYPE.PROTECT:
                    // case BIFFRECORDTYPE.PROT4REVPASSWORD:
                        // IsProtected
                        // break;
                    // case BIFFRECORDTYPE.PASSWORD:
                    default:
                        break;
                }
            }

            if (SST != null)
            {
                SST.Flush();
            }

            foreach (var format in formats)
            {
                // We don't decode the value until here in-case there are format records before the 
                // codepage record. 
                Formats.Add(format.Key, new NumberFormatString(format.Value.GetValue(Encoding)));
            }
        }
    }
}
