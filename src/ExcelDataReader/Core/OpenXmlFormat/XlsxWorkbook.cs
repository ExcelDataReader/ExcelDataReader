using System.Collections.Generic;
using System.Xml;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxWorkbook : IWorkbook<XlsxWorksheet>
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NsRelationship = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string ElementSst = "sst";
        private const string ElementSheets = "sheets";
        private const string ElementSheet = "sheet";
        private const string ElementT = "t";
        private const string ElementR = "r";
        private const string ElementStringItem = "si";
        private const string ElementStyleSheet = "styleSheet";
        private const string ElementCellCrossReference = "cellXfs";
        private const string ElementNumberFormats = "numFmts";
        private const string ElementWorkbook = "workbook";
        private const string ElementWorkbookProperties = "workbookPr";

        private const string AttributeSheetId = "sheetId";
        private const string AttributeVisibleState = "state";
        private const string AttributeName = "name";
        private const string AttributeRelationshipId = "r:id";

        private const string ElementRelationship = "Relationship";
        private const string ElementRelationships = "Relationships";
        private const string AttributeId = "Id";
        private const string AttributeTarget = "Target";

        private const string NXF = "xf";
        private const string ANumFmtId = "numFmtId";
        private const string AXFId = "xfId";
        private const string AApplyNumberFormat = "applyNumberFormat";

        private const string NNumFmt = "numFmt";
        private const string AFormatCode = "formatCode";

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

        public XlsxStyles Styles { get; } = new XlsxStyles();

        public bool IsDate1904 { get; private set; }

        public int ResultsCount => Sheets?.Count ?? -1;

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
                        reader.GetAttribute(AttributeRelationshipId),
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
                    ReadCellXfs(reader);
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

        private void ReadCellXfs(XmlReader reader)
        {
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(NXF, NsSpreadsheetMl))
                {
                    var xfId = reader.GetAttribute(AXFId);
                    var numFmtId = reader.GetAttribute(ANumFmtId);

                    Styles.CellXfs.Add(
                        new XlsxXf(
                            xfId == null ? -1 : int.Parse(xfId),
                            numFmtId == null ? -1 : int.Parse(numFmtId),
                            reader.GetAttribute(AApplyNumberFormat)));
                    reader.Skip();
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
                    Styles.NumFmts.Add(
                        new XlsxNumFmt(
                            int.Parse(reader.GetAttribute(ANumFmtId)),
                            reader.GetAttribute(AFormatCode)));
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
