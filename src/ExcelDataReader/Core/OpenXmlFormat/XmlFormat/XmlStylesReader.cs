using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal sealed class XmlStylesReader : XmlRecordReader
    {
        private const string ElementStyleSheet = "styleSheet";

        private const string ANumFmtId = "numFmtId";

        private const string ElementCellCrossReference = "cellXfs";
        private const string ElementCellStyleCrossReference = "cellStyleXfs";
        private const string NXF = "xf";
        private const string AXFId = "xfId";

        // private const string AApplyNumberFormat = "applyNumberFormat";
        // private const string AApplyAlignment = "applyAlignment";
        // private const string AApplyProtection = "applyProtection";
        private const string ElementNumberFormats = "numFmts";
        private const string NNumFmt = "numFmt";
        private const string AFormatCode = "formatCode";

        private const string NAlignment = "alignment";
        private const string AIndent = "indent";
        private const string AHorizontal = "horizontal";

        private const string NProtection = "protection";
        private const string AHidden = "hidden";
        private const string ALocked = "locked";

        public XmlStylesReader(XmlReader reader) 
            : base(reader)
        {
        }

        protected override IEnumerable<Record> ReadOverride()
        {
            if (!Reader.IsStartElement(ElementStyleSheet, ProperNamespaces.NsSpreadsheetMl))
            {
                yield break;
            }

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(ElementCellCrossReference, ProperNamespaces.NsSpreadsheetMl))
                {
                    foreach (var xf in ReadCellXfs(ProperNamespaces.NsSpreadsheetMl))
                        yield return new ExtendedFormatRecord(xf);
                }
                else if (Reader.IsStartElement(ElementCellStyleCrossReference, ProperNamespaces.NsSpreadsheetMl))
                {
                    foreach (var xf in ReadCellXfs(ProperNamespaces.NsSpreadsheetMl))
                        yield return new CellStyleExtendedFormatRecord(xf);
                }
                else if (Reader.IsStartElement(ElementNumberFormats, ProperNamespaces.NsSpreadsheetMl))
                {
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        continue;
                    }

                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(NNumFmt, ProperNamespaces.NsSpreadsheetMl))
                        {
                            int.TryParse(Reader.GetAttribute(ANumFmtId), NumberStyles.Integer, CultureInfo.InvariantCulture, out var numFmtId);
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

        private IEnumerable<ExtendedFormat> ReadCellXfs(string nsSpreadsheetMl)
        {
            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(NXF, nsSpreadsheetMl))
                {
                    int.TryParse(Reader.GetAttribute(AXFId), NumberStyles.Integer, CultureInfo.InvariantCulture, out var xfId);
                    int.TryParse(Reader.GetAttribute(ANumFmtId), NumberStyles.Integer, CultureInfo.InvariantCulture, out var numFmtId);
                    
                    // var applyNumberFormat = Reader.GetAttribute(AApplyNumberFormat) == "1";
                    // var applyAlignment = Reader.GetAttribute(AApplyAlignment) == "1";
                    // var applyProtection = Reader.GetAttribute(AApplyProtection) == "1";
                    ReadAlignment(Reader, nsSpreadsheetMl, out int indentLevel, out HorizontalAlignment horizontalAlignment, out var hidden, out var locked);

                    yield return new ExtendedFormat(xfId, -1, numFmtId, locked, hidden, indentLevel, horizontalAlignment);

                    // reader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }

            static void ReadAlignment(XmlReader reader, string nsSpreadsheetMl, out int indentLevel, out HorizontalAlignment horizontalAlignment, out bool hidden, out bool locked)
            {
                indentLevel = 0;
                horizontalAlignment = HorizontalAlignment.General;
                hidden = false;
                locked = false;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                {
                    return;
                }

                while (!reader.EOF)
                {
                    if (reader.IsStartElement(NAlignment, nsSpreadsheetMl))
                    {
                        int.TryParse(reader.GetAttribute(AIndent), NumberStyles.Integer, CultureInfo.InvariantCulture, out indentLevel);
                        try
                        {
                            horizontalAlignment = (HorizontalAlignment)Enum.Parse(typeof(HorizontalAlignment), reader.GetAttribute(AHorizontal), true);
                        }
                        catch (ArgumentException)
                        {
                        }
                        catch (OverflowException)
                        {
                        }

                        reader.Skip();
                    }
                    else if (reader.IsStartElement(NProtection, nsSpreadsheetMl))
                    {
                        locked = reader.GetAttribute(ALocked) == "1";
                        hidden = reader.GetAttribute(AHidden) == "1";
                        reader.Skip();
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }
        }
    }
}
