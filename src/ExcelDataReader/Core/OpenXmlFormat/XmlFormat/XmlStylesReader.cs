using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal sealed class XmlStylesReader : XmlRecordReader
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private const string ElementStyleSheet = "styleSheet";

        private const string ANumFmtId = "numFmtId";

        private const string ElementCellCrossReference = "cellXfs";
        private const string NXF = "xf";
        private const string AXFId = "xfId";
        private const string AApplyNumberFormat = "applyNumberFormat";

        private const string ElementNumberFormats = "numFmts";
        private const string NNumFmt = "numFmt";
        private const string AFormatCode = "formatCode";
                
        public XmlStylesReader(XmlReader reader) 
            : base(reader)
        {
        }

        protected override IEnumerable<Record> ReadOverride()
        {
            if (!Reader.IsStartElement(ElementStyleSheet, NsSpreadsheetMl))
            {
                yield break;
            }

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(ElementCellCrossReference, NsSpreadsheetMl))
                {
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        continue;
                    }

                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(NXF, NsSpreadsheetMl))
                        {
                            int.TryParse(Reader.GetAttribute(AXFId), out var xfId);
                            int.TryParse(Reader.GetAttribute(ANumFmtId), out var numFmtId);
                            var applyNumberFormat = Reader.GetAttribute(AApplyNumberFormat) != "0";
                            yield return new ExtendedFormatRecord(xfId, numFmtId, applyNumberFormat);
                            Reader.Skip();
                        }
                        else if (!XmlReaderHelper.SkipContent(Reader))
                        {
                            break;
                        }
                    }
                }
                else if (Reader.IsStartElement(ElementNumberFormats, NsSpreadsheetMl))
                {
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        continue;
                    }

                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(NNumFmt, NsSpreadsheetMl))
                        {
                            int.TryParse(Reader.GetAttribute(ANumFmtId), out var numFmtId);
                            var formatCode = Reader.GetAttribute(AFormatCode);

                            yield return new NumberFormatRecord(numFmtId, formatCode);
                            Reader.Skip();
                        }
                        else if (!XmlReaderHelper.SkipContent(Reader))
                        {
                            break;
                        }
                    }
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }
        }
    }
}
