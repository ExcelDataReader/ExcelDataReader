using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxWorksheet : IWorksheet
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NDimension = "dimension";
        private const string NWorksheet = "worksheet";
        private const string NRow = "row";
        private const string NCol = "col";
        private const string NC = "c"; // cell
        private const string NV = "v";
        private const string NIs = "is";
        private const string NT = "t";
        private const string ARef = "ref";
        private const string AR = "r";
        private const string AT = "t";
        private const string AS = "s";
        private const string NSheetData = "sheetData";
        private const string NMergeCells = "mergeCells";
        private const string NMergeCell = "mergeCell";

        private const string NInlineStr = "inlineStr";
        private const string NStr = "str";

        private const string NHeaderFooter = "headerFooter";
        private const string ADifferentFirst = "differentFirst";
        private const string ADifferentOddEven = "differentOddEven";
        private const string NFirstHeader = "firstHeader";
        private const string NFirstFooter = "firstFooter";
        private const string NOddHeader = "oddHeader";
        private const string NOddFooter = "oddFooter";
        private const string NEvenHeader = "evenHeader";
        private const string NEvenFooter = "evenFooter";

        private const string NSheetProperties = "sheetPr";
        private const string NSheetFormatProperties = "sheetFormatPr";
        private const string ADefaultRowHeight = "defaultRowHeight";
        private const string AHidden = "hidden";
        private const string ACustomHeight = "customHeight";
        private const string AHt = "ht";

        public XlsxWorksheet(ZipWorker document, XlsxWorkbook workbook, XlsxBoundSheet refSheet)
        {
            Document = document;
            Workbook = workbook;

            Name = refSheet.Name;
            Id = refSheet.Id;
            Rid = refSheet.Rid;
            VisibleState = refSheet.VisibleState;
            Path = refSheet.Path;
            DefaultRowHeight = 12.75; // 255 twips

            ReadWorksheetGlobals();
        }

        public int FieldCount { get; private set; }

        public int RowCount { get; private set; }

        public string Name { get; }

        public string CodeName { get; private set; }

        public string VisibleState { get; }

        public HeaderFooter HeaderFooter { get; private set; }

        public double DefaultRowHeight { get; private set; }

        public int Id { get; }

        public string Rid { get; set; }

        public string Path { get; set; }

        public CellRange[] MergeCells { get; private set; }

        private ZipWorker Document { get; }

        private XlsxWorkbook Workbook { get; }

        public IEnumerable<Row> ReadRows()
        {
            var rowIndex = 0;
            foreach (var sheetObject in ReadWorksheetStream(false))
            {
                if (sheetObject.Type == XlsxElementType.Row)
                {
                    var rowBlock = ((XlsxRow)sheetObject).Row;

                    for (; rowIndex < rowBlock.RowIndex; ++rowIndex)
                    {
                        yield return new Row()
                        {
                            RowIndex = rowIndex,
                            Height = DefaultRowHeight,
                            Cells = new List<Cell>()
                        };
                    }

                    rowIndex++;
                    yield return rowBlock;
                }
            }
        }

        public NumberFormatString GetNumberFormatString(int numberFormatIndex)
        {
            var numFmt = Workbook.Styles.NumFmts.Find(x => x.Id == numberFormatIndex);
            if (numFmt != null)
            { 
                return numFmt.FormatCode;
            }

            return BuiltinNumberFormat.GetBuiltinNumberFormat(numberFormatIndex);
        }

        private void ReadWorksheetGlobals()
        {
            if (string.IsNullOrEmpty(Path))
                return;

            int rows = int.MinValue;
            int cols = int.MinValue;
            foreach (var sheetObject in ReadWorksheetStream(false))
            {
                if (sheetObject.Type == XlsxElementType.Row)
                {
                    var rowBlock = ((XlsxRow)sheetObject).Row;
                    rows = Math.Max(rows, rowBlock.RowIndex);
                    cols = Math.Max(cols, rowBlock.GetMaxColumnIndex());
                }
                else if (sheetObject.Type == XlsxElementType.HeaderFooter)
                {
                    XlsxHeaderFooter headerFooter = (XlsxHeaderFooter)sheetObject;
                    HeaderFooter = headerFooter.Value;
                }
                else if (sheetObject.Type == XlsxElementType.MergeCells)
                {
                    XlsxMergeCells mergeCells = (XlsxMergeCells)sheetObject;
                    MergeCells = mergeCells.Value.ToArray();
                }
            }

            if (rows != int.MinValue && cols != int.MinValue)
            {
                FieldCount = cols + 1;
                RowCount = rows + 1;
            }
        }

        private IEnumerable<XlsxElement> ReadWorksheetStream(bool skipSheetData)
        {
            if (string.IsNullOrEmpty(Path))
                yield break;

            using (var sheetStream = Document.GetWorksheetStream(Path))
            {
                if (sheetStream == null)
                {
                    yield break;
                }

                using (var xmlReader = XmlReader.Create(sheetStream))
                {
                    foreach (var sheetObject in ReadWorksheetStream(xmlReader, skipSheetData))
                    {
                        yield return sheetObject;
                    }
                }
            }
        }

        private IEnumerable<XlsxElement> ReadWorksheetStream(XmlReader xmlReader, bool skipSheetData)
        {
            if (!xmlReader.IsStartElement(NWorksheet, NsSpreadsheetMl))
            {
                yield break;
            }

            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                yield break;
            }

            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NDimension, NsSpreadsheetMl))
                {
                    var dimension = ReadDimension(xmlReader);
                    if (dimension != null)
                        yield return dimension;
                }
                else if (xmlReader.IsStartElement(NSheetData, NsSpreadsheetMl))
                {
                    if (skipSheetData)
                    {
                        xmlReader.Skip();
                    }
                    else
                    {
                        foreach (var row in ReadSheetData(xmlReader))
                        {
                            yield return row;
                        }
                    }
                }
                else if (xmlReader.IsStartElement(NMergeCells, NsSpreadsheetMl))
                {
                    var result = ReadMergeCells(xmlReader);
                    if (result != null)
                        yield return result;
                }
                else if (xmlReader.IsStartElement(NHeaderFooter, NsSpreadsheetMl))
                {
                    var result = ReadHeaderFooter(xmlReader);
                    if (result != null)
                        yield return result;
                }
                else if (xmlReader.IsStartElement(NSheetProperties, NsSpreadsheetMl))
                {
                    var codeName = xmlReader.GetAttribute("codeName");
                    if (!string.IsNullOrEmpty(codeName))
                        CodeName = codeName;

                    xmlReader.Skip();
                }
                else if (xmlReader.IsStartElement(NSheetFormatProperties, NsSpreadsheetMl))
                {
                    if (double.TryParse(xmlReader.GetAttribute(ADefaultRowHeight), NumberStyles.Any, CultureInfo.InvariantCulture, out var defaultRowHeight))
                        DefaultRowHeight = defaultRowHeight;

                    xmlReader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }
        }

        private XlsxDimension ReadDimension(XmlReader xmlReader)
        {
            var dimValue = xmlReader.GetAttribute(ARef);
            xmlReader.Skip();

            if (!string.IsNullOrEmpty(dimValue))
            {
                var dimension = new XlsxDimension(dimValue);
                if (dimension.IsRange)
                {
                    return dimension;
                }
            }

            return null;
        }

        private IEnumerable<XlsxRow> ReadSheetData(XmlReader xmlReader)
        {
            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                yield break;
            }

            int nextRowIndex = 0;
            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NRow, NsSpreadsheetMl))
                {
                    var row = ReadRow(xmlReader, nextRowIndex);
                    nextRowIndex = row.RowIndex + 1;
                    yield return new XlsxRow()
                    {
                        Row = row
                    };
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }
        }

        private XlsxMergeCells ReadMergeCells(XmlReader xmlReader)
        {
            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                return null;
            }

            var ranges = new List<CellRange>();

            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NMergeCell, NsSpreadsheetMl))
                {
                    var cellRefs = xmlReader.GetAttribute(ARef);
                    string from = string.Empty, to = string.Empty;
                    var fromTo = cellRefs.Split(':');

                    if (fromTo.Length == 2)
                    {
                        from = fromTo[0];
                        to = fromTo[1];
                    }

                    ranges.Add(new CellRange(from, to));

                    xmlReader.Read();
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }

            return new XlsxMergeCells()
            {
                Value = ranges
            };
        }

        private XlsxHeaderFooter ReadHeaderFooter(XmlReader xmlReader)
        {
            var differentFirst = xmlReader.GetAttribute(ADifferentFirst) == "1";
            var differentOddEven = xmlReader.GetAttribute(ADifferentOddEven) == "1";

            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                return null;
            }

            var headerFooter = new HeaderFooter(differentFirst, differentOddEven);

            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NOddHeader, NsSpreadsheetMl))
                {
                    headerFooter.OddHeader = xmlReader.ReadElementContentAsString();
                }
                else if (xmlReader.IsStartElement(NOddFooter, NsSpreadsheetMl))
                {
                    headerFooter.OddFooter = xmlReader.ReadElementContentAsString();
                }
                else if (xmlReader.IsStartElement(NEvenHeader, NsSpreadsheetMl))
                {
                    headerFooter.EvenHeader = xmlReader.ReadElementContentAsString();
                }
                else if (xmlReader.IsStartElement(NEvenFooter, NsSpreadsheetMl))
                {
                    headerFooter.EvenFooter = xmlReader.ReadElementContentAsString();
                }
                else if (xmlReader.IsStartElement(NFirstHeader, NsSpreadsheetMl))
                {
                    headerFooter.FirstHeader = xmlReader.ReadElementContentAsString();
                }
                else if (xmlReader.IsStartElement(NFirstFooter, NsSpreadsheetMl))
                {
                    headerFooter.FirstFooter = xmlReader.ReadElementContentAsString();
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }

            return new XlsxHeaderFooter(headerFooter);
        }

        private Row ReadRow(XmlReader xmlReader, int nextRowIndex)
        {
            var result = new Row()
            {
                Cells = new List<Cell>()
            };

            if (int.TryParse(xmlReader.GetAttribute(AR), out int rowIndex))
                result.RowIndex = rowIndex - 1; // The row attribute is 1-based
            else
                result.RowIndex = nextRowIndex;

            int.TryParse(xmlReader.GetAttribute(AHidden), out int hidden);
            int.TryParse(xmlReader.GetAttribute(ACustomHeight), out int customHeight);
            double.TryParse(xmlReader.GetAttribute(AHt), NumberStyles.Any, CultureInfo.InvariantCulture, out var height);

            if (hidden == 0)
                result.Height = customHeight != 0 ? height : DefaultRowHeight;

            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                return result;
            }

            int nextColumnIndex = 0;
            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NC, NsSpreadsheetMl))
                {
                    var cell = ReadCell(xmlReader, nextColumnIndex);
                    nextColumnIndex = cell.ColumnIndex + 1;
                    result.Cells.Add(cell);
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }

            return result;
        }

        private Cell ReadCell(XmlReader xmlReader, int nextColumnIndex)
        {
            var result = new Cell();

            var aS = xmlReader.GetAttribute(AS);
            var aT = xmlReader.GetAttribute(AT);
            var aR = xmlReader.GetAttribute(AR);

            if (ReferenceHelper.ParseReference(aR, out int referenceColumn, out int referenceRow))
                result.ColumnIndex = referenceColumn - 1; // ParseReference is 1-based
            else
                result.ColumnIndex = nextColumnIndex;

            if (aS != null)
            {
                if (int.TryParse(aS, NumberStyles.Any, CultureInfo.InvariantCulture, out var styleIndex))
                {
                    if (styleIndex >= 0 && styleIndex < Workbook.Styles.CellXfs.Count)
                    {
                        XlsxXf xf = Workbook.Styles.CellXfs[styleIndex];
                        result.NumberFormatIndex = xf.NumFmtId;
                    }
                }
            }

            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                return result;
            }

            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NV, NsSpreadsheetMl))
                {
                    var rawValue = xmlReader.ReadElementContentAsString();
                    if (!string.IsNullOrEmpty(rawValue))
                        result.Value = ConvertCellValue(rawValue, aT, result.NumberFormatIndex);
                }
                else if (xmlReader.IsStartElement(NIs, NsSpreadsheetMl))
                {
                    var rawValue = XlsxWorkbook.ReadStringItem(xmlReader);
                    if (!string.IsNullOrEmpty(rawValue))
                        result.Value = ConvertCellValue(rawValue, aT, result.NumberFormatIndex);
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }

            return result;
        }

        private object ConvertCellValue(string rawValue, string aT, int numberFormatIndex)
        {
            const NumberStyles style = NumberStyles.Any;
            var invariantCulture = CultureInfo.InvariantCulture;

            switch (aT)
            {
                case AS: //// if string
                    if (int.TryParse(rawValue, style, invariantCulture, out var sstIndex))
                    {
                        if (sstIndex >= 0 && sstIndex < Workbook.SST.Count)
                        {
                            return Helpers.ConvertEscapeChars(Workbook.SST[sstIndex]);
                        }
                    }

                    return rawValue;
                case NInlineStr: //// if string inline
                case NStr: //// if cached formula string
                    return Helpers.ConvertEscapeChars(rawValue);
                case "b": //// boolean
                    return rawValue == "1";
                case "d": //// ISO 8601 date
                    if (DateTime.TryParseExact(rawValue, "yyyy-MM-dd", invariantCulture, DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite, out var date))
                        return date;

                    return rawValue;
                default:
                    if (double.TryParse(rawValue, style, invariantCulture, out double number))
                    {
                        var format = GetNumberFormatString(numberFormatIndex);
                        if (format != null)
                        {
                            if (format.IsDateTimeFormat)
                                return Helpers.ConvertFromOATime(number, Workbook.IsDate1904);
                            if (format.IsTimeSpanFormat)
                                return TimeSpan.FromDays(number);
                        }

                        return number;
                    }

                    return rawValue;
            }
        }
    }
}
