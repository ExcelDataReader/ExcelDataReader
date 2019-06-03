using System;
using System.Collections.Generic;
using System.Xml;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxWorkbook : CommonWorkbook, IWorkbook<XlsxWorksheet>
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NsRelationship = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string NsDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string ElementSst = "sst";
        private const string ElementSheets = "sheets";
        private const string ElementSheet = "sheet";
        private const string ElementT = "t";
        private const string ElementR = "r";
        private const string ElementStringItem = "si";
        private const string ElementStyleSheet = "styleSheet";
        private const string ElementCellCrossReference = "cellXfs";
        private const string ElementCellStyleCrossReference = "cellStyleXfs";
        private const string ElementNumberFormats = "numFmts";
        private const string ElementWorkbook = "workbook";
        private const string ElementWorkbookProperties = "workbookPr";

        private const string AttributeSheetId = "sheetId";
        private const string AttributeVisibleState = "state";
        private const string AttributeName = "name";
        private const string AttributeRelationshipId = "id";

        private const string ElementRelationship = "Relationship";
        private const string ElementRelationships = "Relationships";
        private const string AttributeId = "Id";
        private const string AttributeTarget = "Target";

        private const string NXF = "xf";
        private const string ANumFmtId = "numFmtId";
        private const string AXFId = "xfId";
        private const string AApplyNumberFormat = "applyNumberFormat";
        private const string AApplyAlignment = "applyAlignment";

        private const string NNumFmt = "numFmt";
        private const string AFormatCode = "formatCode";

        private const string NAlignment = "alignment";
        private const string AIndent = "indent";
        private const string AHorizontal = "horizontal";

        private ZipWorker _zipWorker;

        public XlsxWorkbook(ZipWorker zipWorker)
        {
            _zipWorker = zipWorker;

            ReadWorkbook();
            ReadWorkbookRels();
            ReadSharedStrings();
            ReadStyles();
        }

        public List<XlsxBoundSheet> Sheets { get; } = new List<XlsxBoundSheet>();

        public XlsxSST SST { get; } = new XlsxSST();

        public bool IsDate1904 { get; private set; }

        public int ResultsCount => Sheets?.Count ?? -1;

        private List<ExtendedFormat> CellFormats { get; } = new List<ExtendedFormat>();

        private List<ExtendedFormat> CellStyleFormats { get; } = new List<ExtendedFormat>();

        public static string ReadStringItem(XmlReader reader)
        {
            string result = string.Empty;
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return result;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementT, NsSpreadsheetMl))
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    result += reader.ReadElementContentAsString();
                }
                else if (reader.IsStartElement(ElementR, NsSpreadsheetMl))
                {
                    result += ReadRichTextRun(reader);
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return result;
        }

        public IEnumerable<XlsxWorksheet> ReadWorksheets()
        {
            foreach (var sheet in Sheets)
            {
                yield return new XlsxWorksheet(_zipWorker, this, sheet);
            }
        }

        public void GetCellStyleFromXF(CellStyle cellStyle, int styleIndex)
        {
            if (styleIndex < 0 || styleIndex >= CellFormats.Count)
            {
                // Invalid index, return default.
                return;
            }

            var extendedFormat = CellFormats[styleIndex];

            cellStyle.IndentLevel = extendedFormat.IndentLevel;
            cellStyle.HorizontalAlignment = extendedFormat.HorizontalAlignment;
            cellStyle.FormatIndex = GetNumberFormatFromFileIndex(extendedFormat.FormatIndex);

            if (extendedFormat.XfId >= 0 && extendedFormat.XfId < CellStyleFormats.Count)
            {
                extendedFormat = CellStyleFormats[extendedFormat.XfId];

                if (extendedFormat.ApplyNumberFormat)
                    cellStyle.FormatIndex = GetNumberFormatFromFileIndex(extendedFormat.FormatIndex);
                if (extendedFormat.ApplyAlignment)
                {
                    cellStyle.IndentLevel = extendedFormat.IndentLevel;
                    cellStyle.HorizontalAlignment = extendedFormat.HorizontalAlignment;
                }
            }
        }

        private static string ReadRichTextRun(XmlReader reader)
        {
            string result = string.Empty;
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return result;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementT, NsSpreadsheetMl))
                {
                    result += reader.ReadElementContentAsString();
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return result;
        }

        private static IEnumerable<ExtendedFormat> ReadCellXfs(XmlReader reader)
        {
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                yield break;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(NXF, NsSpreadsheetMl))
                {
                    // As far as I can tell the default is false / "0" for the apply attributes. 
                    var applyNumberFormat = reader.GetAttribute(AApplyNumberFormat) == "1";
                    var applyAlignment = reader.GetAttribute(AApplyAlignment) == "1";

                    int.TryParse(reader.GetAttribute(AXFId), out var xfId);
                    int.TryParse(reader.GetAttribute(ANumFmtId), out var numFmtId);

                    ReadAlignment(reader, out int indentLevel, out HorizontalAlignment horizontalAlignment);

                    yield return new ExtendedFormat
                    {
                        XfId = xfId,
                        ApplyNumberFormat = applyNumberFormat,
                        FormatIndex = numFmtId,
                        ApplyAlignment = applyAlignment,
                        IndentLevel = indentLevel,
                        HorizontalAlignment = horizontalAlignment,
                    };
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    yield break;
                }
            }
        }

        private static void ReadAlignment(XmlReader reader, out int indentLevel, out HorizontalAlignment horizontalAlignment)
        {
            indentLevel = 0;
            horizontalAlignment = HorizontalAlignment.General;

            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(NAlignment, NsSpreadsheetMl))
                {
                    int.TryParse(reader.GetAttribute(AIndent), out indentLevel);
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
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }
        }

        private void ReadWorkbook()
        {
            using (var stream = _zipWorker.GetWorkbookStream())
            {
                if (stream == null)
                {
                    throw new Exceptions.HeaderException(Errors.ErrorZipNoOpenXml);
                }

                using (XmlReader reader = XmlReader.Create(stream))
                {
                    ReadWorkbook(reader);
                }
            }
        }

        private void ReadWorkbook(XmlReader reader)
        {
            if (!reader.IsStartElement(ElementWorkbook, NsSpreadsheetMl))
            {
                return;
            }

            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementWorkbookProperties, NsSpreadsheetMl))
                {
                    // Workbook VBA CodeName: reader.GetAttribute("codeName");
                    IsDate1904 = reader.GetAttribute("date1904") == "1";
                    reader.Skip();
                }
                else if (reader.IsStartElement(ElementSheets, NsSpreadsheetMl))
                {
                    ReadSheets(reader);
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }
        }

        private void ReadSheets(XmlReader reader)
        {
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementSheet, NsSpreadsheetMl))
                {
                    Sheets.Add(new XlsxBoundSheet(
                        reader.GetAttribute(AttributeName),
                        int.Parse(reader.GetAttribute(AttributeSheetId)),
                        reader.GetAttribute(AttributeRelationshipId, NsDocumentRelationship),
                        reader.GetAttribute(AttributeVisibleState)));
                    reader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }
        }

        private void ReadWorkbookRels()
        {
            using (var stream = _zipWorker.GetWorkbookRelsStream())
            {
                if (stream == null)
                {
                    return;
                }

                using (XmlReader reader = XmlReader.Create(stream))
                {
                    ReadWorkbookRels(reader);
                }
            }
        }

        private void ReadWorkbookRels(XmlReader reader)
        {
            if (!reader.IsStartElement(ElementRelationships, NsRelationship))
            {
                return;
            }

            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementRelationship, NsRelationship))
                {
                    string rid = reader.GetAttribute(AttributeId);
                    foreach (var sheet in Sheets)
                    {
                        if (sheet.Rid == rid)
                        {
                            sheet.Path = reader.GetAttribute(AttributeTarget);
                            break;
                        }
                    }

                    reader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }
        }

        private void ReadSharedStrings()
        {
            using (var stream = _zipWorker.GetSharedStringsStream())
            {
                if (stream == null)
                    return;

                using (XmlReader reader = XmlReader.Create(stream))
                {
                    ReadSharedStrings(reader);
                }
            }
        }

        private void ReadSharedStrings(XmlReader reader)
        {
            if (!reader.IsStartElement(ElementSst, NsSpreadsheetMl))
            {
                return;
            }

            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementStringItem, NsSpreadsheetMl))
                {
                    var value = ReadStringItem(reader);
                    SST.Add(value);
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }
        }

        private void ReadStyles()
        {
            using (var stream = _zipWorker.GetStylesStream())
            {
                if (stream == null)
                    return;

                using (XmlReader reader = XmlReader.Create(stream))
                {
                    ReadStyles(reader);
                }
            }
        }

        private void ReadStyles(XmlReader reader)
        {
            if (!reader.IsStartElement(ElementStyleSheet, NsSpreadsheetMl))
            {
                return;
            }

            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementCellCrossReference, NsSpreadsheetMl))
                {
                    foreach (var xf in ReadCellXfs(reader))
                    {
                        CellFormats.Add(xf);
                    }
                }
                else if (reader.IsStartElement(ElementCellStyleCrossReference, NsSpreadsheetMl))
                {
                    foreach (var xf in ReadCellXfs(reader))
                    {
                        CellStyleFormats.Add(xf);
                    }
                }
                else if (reader.IsStartElement(ElementNumberFormats, NsSpreadsheetMl))
                {
                    ReadNumberFormats(reader);
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }
        }

        private void ReadNumberFormats(XmlReader reader)
        {
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(NNumFmt, NsSpreadsheetMl))
                {
                    int.TryParse(reader.GetAttribute(ANumFmtId), out var numFmtId);
                    var formatCode = reader.GetAttribute(AFormatCode);

                    AddNumberFormat(numFmtId, formatCode);
                    reader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }
        }

        private sealed class ExtendedFormat
        {
            public int XfId { get; set; }

            public bool ApplyNumberFormat { get; set; }

            public int FormatIndex { get; set; }

            public bool ApplyAlignment { get; set; }

            public int IndentLevel { get; set; }

            public HorizontalAlignment HorizontalAlignment { get; set; }
        }
    }
}
