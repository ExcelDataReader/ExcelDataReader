// ReSharper disable InconsistentNaming
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

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

        public XlsxWorksheet(ZipWorker document, XlsxWorkbook workbook, XlsxBoundSheet refSheet)
        {
            Document = document;
            Workbook = workbook;

            Name = refSheet.Name;
            Id = refSheet.Id;
            Rid = refSheet.Rid;
            VisibleState = refSheet.VisibleState;
            Path = refSheet.Path;

            ReadWorksheetGlobals();
        }

        public XlsxDimension Dimension { get; set; }

        public int ColumnsCount => Dimension?.LastCol ?? 0;

        public int FieldCount => ColumnsCount;

        public int RowsCount => Dimension == null ? -1 : Dimension.LastRow - Dimension.FirstRow + 1;

        public string Name { get; }

        public string CodeName { get; private set; }

        public string VisibleState { get; }

        public HeaderFooter HeaderFooter { get; private set; }

        public int Id { get; }

        public string Rid { get; set; }

        public string Path { get; set; }

        private ZipWorker Document { get; }

        private XlsxWorkbook Workbook { get; }

        public IEnumerable<object[]> ReadRows()
        {
            if (Dimension == null)
            {
                yield break;
            }

            var rowIndex = 1;
            foreach (var sheetObject in ReadWorksheetStream(false))
            {
                if (sheetObject.Type == XlsxElementType.Row)
                {
                    var rowBlock = (XlsxRow)sheetObject;

                    for (; rowIndex < rowBlock.RowIndex; ++rowIndex)
                    {
                        yield return new object[FieldCount];
                    }

                    rowIndex++;
                    var result = new object[FieldCount];
                    foreach (var cell in rowBlock.Cells)
                    {
                        var columnIndex = cell.ColumnIndex - 1; // from 1 to 0-based
                        if (columnIndex < result.Length)
                            result[columnIndex] = cell.Value;
                    }

                    yield return result;
                }
            }
        }

        private void ReadWorksheetGlobals()
        {
            if (string.IsNullOrEmpty(Path))
                return;

            foreach (var sheetObject in ReadWorksheetStream(true))
            {
                switch (sheetObject.Type)
                {
                    case XlsxElementType.Dimension:
                        Dimension = (XlsxDimension)sheetObject;
                        break;
                    case XlsxElementType.HeaderFooter:
                        XlsxHeaderFooter headerFooter = (XlsxHeaderFooter)sheetObject;
                        HeaderFooter = headerFooter.Value;
                        break;
                }
            }
            
            if (Dimension == null)
            {
                int rows = int.MinValue;
                int cols = int.MinValue;
                foreach (var sheetObject in ReadWorksheetStream(false))
                {
                    if (sheetObject.Type == XlsxElementType.Row)
                    {
                        var rowBlock = (XlsxRow)sheetObject;
                        rows = Math.Max(rows, rowBlock.RowIndex);
                        cols = Math.Max(cols, rowBlock.GetMaxColumnIndex());
                    }
                }

                if (rows != int.MinValue && cols != int.MinValue)
                {
                    Dimension = new XlsxDimension(rows, cols);
                }
            }
        }

        private IEnumerable<XlsxElement> ReadWorksheetStream(bool skipSheetData)
        {
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

            int nextRowIndex = 1;
            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NRow, NsSpreadsheetMl))
                {
                    var row = ReadRow(xmlReader, nextRowIndex);
                    nextRowIndex = row.RowIndex + 1;
                    yield return row;
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }
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

        private XlsxRow ReadRow(XmlReader xmlReader, int nextRowIndex)
        {
            var result = new XlsxRow();

            if (int.TryParse(xmlReader.GetAttribute(AR), out int rowIndex))
                result.RowIndex = rowIndex;
            else
                result.RowIndex = nextRowIndex;

            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                return result;
            }

            int nextColumnIndex = 1;
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

        private XlsxCell ReadCell(XmlReader xmlReader, int nextColumnIndex)
        {
            var result = new XlsxCell();

            var aS = xmlReader.GetAttribute(AS);
            var aT = xmlReader.GetAttribute(AT);
            var aR = xmlReader.GetAttribute(AR);

            if (ReferenceHelper.ParseReference(aR, out int referenceColumn, out int referenceRow))
                result.ColumnIndex = referenceColumn;
            else
                result.ColumnIndex = nextColumnIndex;

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
                        result.Value = ConvertCellValue(rawValue, aT, aS);
                }
                else if (xmlReader.IsStartElement(NIs, NsSpreadsheetMl))
                {
                    var rawValue = ReadInlineString(xmlReader);
                    if (!string.IsNullOrEmpty(rawValue))
                        result.Value = ConvertCellValue(rawValue, aT, aS);
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }

            return result;
        }

        private string ReadInlineString(XmlReader xmlReader)
        {
            string result = null;

            if (!XmlReaderHelper.ReadFirstContent(xmlReader))
            {
                return result;
            }

            while (!xmlReader.EOF)
            {
                if (xmlReader.IsStartElement(NT, NsSpreadsheetMl))
                {
                    result = xmlReader.ReadElementContentAsString();
                }
                else if (!XmlReaderHelper.SkipContent(xmlReader))
                {
                    break;
                }
            }

            return result;
        }

        private object ConvertCellValue(string rawValue, string aT, string aS)
        {
            const NumberStyles style = NumberStyles.Any;
            var invariantCulture = CultureInfo.InvariantCulture;

            switch (aT)
            {
                case AS: //// if string
                    return Helpers.ConvertEscapeChars(Workbook.SST[int.Parse(rawValue, invariantCulture)]);
                case NInlineStr: //// if string inline
                case NStr: //// if cached formula string
                    return Helpers.ConvertEscapeChars(rawValue);
                case "b": //// boolean
                    return rawValue == "1";
                case "d": //// ISO 8601 date
                    return DateTime.ParseExact(rawValue, "yyyy-MM-dd", invariantCulture, DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite);
                default:
                    bool isNumber = double.TryParse(rawValue, style, invariantCulture, out double number);

                    if (aS != null)
                    {
                        XlsxXf xf = Workbook.Styles.CellXfs[int.Parse(aS)];
                        if (isNumber && Workbook.IsDateTimeStyle(xf.NumFmtId))
                            return Helpers.ConvertFromOATime(number, Workbook.IsDate1904);

                        // NOTE: Commented out to match behavior of the binary reader; 
                        // formatting should ultimately be applied by the caller
                        // if (xf.NumFmtId == 49) // Text format but value is not stored as a string. If numeric convert to current culture. 
                        //    return isNumber ? number.ToString() : rawValue;
                    }

                    if (isNumber)
                        return number;
                    return rawValue;
            }
        }
    }
}
