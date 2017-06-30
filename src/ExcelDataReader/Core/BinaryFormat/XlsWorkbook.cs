using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader.Exceptions;
using ExcelDataReader.Log;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Globals section of workbook
    /// </summary>
    internal class XlsWorkbook : IWorkbook<XlsWorksheet>
    {
        internal XlsWorkbook(byte[] bytes, bool convertOaDate, Encoding fallbackEncoding)
        {
            Version = 0x0600;
            BiffStream = new XlsBiffStream(bytes);
            ConvertOaDate = convertOaDate;
            ReadWorkbookGlobals(fallbackEncoding);
        }

        public ushort Version { get; set; }

        public Encoding Encoding { get; set; }

        public XlsBiffInterfaceHdr InterfaceHdr { get; set; }

        public XlsBiffRecord Mms { get; set; }

        public XlsBiffRecord WriteAccess { get; set; }

        public XlsBiffSimpleValueRecord CodePage { get; set; }

        public XlsBiffRecord Dsf { get; set; }

        public XlsBiffRecord Country { get; set; }

        public XlsBiffSimpleValueRecord Backup { get; set; }

        public List<XlsBiffRecord> Fonts { get; } = new List<XlsBiffRecord>();

        public Dictionary<ushort, XlsBiffFormatString> Formats { get; } = new Dictionary<ushort, XlsBiffFormatString>();

        public List<XlsBiffRecord> ExtendedFormats { get; } = new List<XlsBiffRecord>();

        public List<XlsBiffRecord> Styles { get; } = new List<XlsBiffRecord>();

        public List<XlsBiffBoundSheet> Sheets { get; } = new List<XlsBiffBoundSheet>();

        /// <summary>
        /// Gets or sets the Shared String Table of workbook
        /// </summary>
        public XlsBiffSST SST { get; set; }

        public XlsBiffRecord ExtSST { get; set; }

        public bool ConvertOaDate { get; }

        public XlsBiffStream BiffStream { get; }

        public bool IsDate1904 { get; private set; }

        public int ResultsCount => Sheets?.Count ?? -1;

        public IEnumerable<XlsWorksheet> ReadWorksheets()
        {
            for (var i = 0; i < Sheets.Count; ++i)
            {
                yield return ReadWorksheet(i);
            }
        }

        public XlsWorksheet ReadWorksheet(int index)
        {
            return new XlsWorksheet(this, index);
        }

        private void ReadWorkbookGlobals(Encoding fallbackEncoding)
        {
            BiffStream.Seek(0, SeekOrigin.Begin);

            XlsBiffRecord rec = BiffStream.Read();
            XlsBiffBOF bof = rec as XlsBiffBOF;

            if (bof == null || bof.Type != BIFFTYPE.WorkbookGlobals)
            {
                throw new ExcelReaderException(Errors.ErrorWorkbookGlobalsInvalidData);
            }

            bool sst = false;

            Version = bof.Version;
            Encoding = BiffStream.BiffVersion == 8 ? Encoding.Unicode : fallbackEncoding;

            while ((rec = BiffStream.Read()) != null)
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
                            Formats.Add((ushort)Formats.Count, fmt);
                        }

                        break;
                    case BIFFRECORDTYPE.FORMAT:
                        {
                            var fmt = (XlsBiffFormatString)rec;
                            Formats.Add(fmt.Index, fmt);
                        }

                        break;
                    case BIFFRECORDTYPE.XF:
                    case BIFFRECORDTYPE.XF_V4:
                    case BIFFRECORDTYPE.XF_V3:
                    case BIFFRECORDTYPE.XF_V2:
                        ExtendedFormats.Add(rec);
                        break;
                    case BIFFRECORDTYPE.SST:
                        SST = (XlsBiffSST)rec;
                        sst = true;
                        break;
                    case BIFFRECORDTYPE.CONTINUE:
                        if (!sst)
                            break;
                        XlsBiffContinue contSST = (XlsBiffContinue)rec;
                        SST.Append(contSST);
                        break;
                    case BIFFRECORDTYPE.EXTSST:
                        ExtSST = rec;
                        sst = false;
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
                    case BIFFRECORDTYPE.EOF:
                        SST?.ReadStrings();
                        return;

                    default:
                        continue;
                }
            }
        }
    }
}