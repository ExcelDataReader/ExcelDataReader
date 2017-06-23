// ReSharper disable InconsistentNaming
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxWorksheet : IWorksheet
    {
        public const string NDimension = "dimension";
        public const string NWorksheet = "worksheet";
        public const string NRow = "row";
        public const string NCol = "col";
        public const string NC = "c"; // cell
        public const string NV = "v";
        public const string NT = "t";
        public const string ARef = "ref";
        public const string AR = "r";
        public const string AT = "t";
        public const string AS = "s";
        public const string NSheetData = "sheetData";
        public const string NInlineStr = "inlineStr";

        private string _namespaceUri;

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

        public bool IsEmpty { get; set; }

        public XlsxDimension Dimension { get; set; }

        public int ColumnsCount => IsEmpty ? 0 : (Dimension?.LastCol ?? -1);

        public int FieldCount => ColumnsCount;

        public int RowsCount => Dimension == null ? -1 : Dimension.LastRow - Dimension.FirstRow + 1;

        public string Name { get; }

        public string VisibleState { get; }

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

            using (var sheetStream = Document.GetWorksheetStream(Path))
            {
                if (sheetStream == null) { 
                    yield break;
                }

                using (var xmlReader = XmlReader.Create(sheetStream))
                {
                    var rowIndex = 0;

                    while (true)
                    { 
                        var rowBlock = ReadSheetRow(xmlReader, this, rowIndex);
                        if (rowBlock == null)
                        {
                            yield break;
                        }

                        for (; rowIndex < rowBlock.RowIndex; ++rowIndex)
                        {
                            yield return new object[FieldCount];
                        }

                        rowIndex++;
                        yield return rowBlock.Values;
                    }
                }
            }
        }

        private XlsxRowBlock ReadSheetRow(XmlReader xmlReader, XlsxWorksheet sheet, int depth)
        {
            var result = new XlsxRowBlock();

            if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.LocalName == XlsxWorksheet.NRow ||
                xmlReader.ReadToFollowing(XlsxWorksheet.NRow, _namespaceUri))
            {
                result.Values = new object[sheet.ColumnsCount];

                int rowIndex; // 1-based
                if (!int.TryParse(xmlReader.GetAttribute(XlsxWorksheet.AR), out rowIndex))
                    rowIndex = depth + 1;

                result.RowIndex = rowIndex - 1; // 0-based

                bool hasValue = false;
                string aS = string.Empty;
                string aT = string.Empty;
                string aR = string.Empty;
                int col = 0;

                while (xmlReader.Read())
                {
                    if (xmlReader.Depth == 2)
                        break;

                    if (xmlReader.NodeType == XmlNodeType.Element)
                    {
                        hasValue = false;

                        if (xmlReader.LocalName == XlsxWorksheet.NC)
                        {
                            aS = xmlReader.GetAttribute(XlsxWorksheet.AS);
                            aT = xmlReader.GetAttribute(XlsxWorksheet.AT);
                            aR = xmlReader.GetAttribute(XlsxWorksheet.AR);

                            if (aR != null)
                            {
                                ReferenceHelper.ParseReference(aR, out col);
                            }
                            else
                            {
                                ++col;
                            }
                        }
                        else if (xmlReader.LocalName == XlsxWorksheet.NV || xmlReader.LocalName == XlsxWorksheet.NT)
                        {
                            hasValue = true;
                        }
                    }

                    if (xmlReader.NodeType == XmlNodeType.Text && hasValue)
                    {
                        if (col - 1 < result.Values.Length)
                            result.Values[col - 1] = ConvertCellValue(xmlReader.Value, aT, aS);
                    }
                }

                return result;
            }

            // not a row
            return null;
        }

        private object ConvertCellValue(string rawValue, string aT, string aS)
        {
            const NumberStyles style = NumberStyles.Any;
            var invariantCulture = CultureInfo.InvariantCulture;

            switch (aT)
            {
                case XlsxWorksheet.AS: //// if string
                    return Helpers.ConvertEscapeChars(Workbook.SST[int.Parse(rawValue, invariantCulture)]);
                case XlsxWorksheet.NInlineStr: //// if string inline
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
                            return Helpers.ConvertFromOATime(number);

                        if (xf.NumFmtId == 49) // Text format but value is not stored as a string. If numeric convert to current culture. 
                            return isNumber ? number.ToString(CultureInfo.CurrentCulture) : rawValue;
                    }

                    if (isNumber)
                        return number;
                    return rawValue;
            }
        }

        private void ReadWorksheetGlobals()
        {
            using (var sheetStream = Document.GetWorksheetStream(Path))
            {
                if (sheetStream == null)
                    return;

                using (var xmlReader = XmlReader.Create(sheetStream))
                {
                    // count rows and cols in case there is no dimension elements
                    int rows = 0;
                    int cols = 0;

                    bool foundDimension = false;

                    _namespaceUri = null;
                    int biggestColumn = 0; // used when no col elements and no dimension
                    int cellElementsInRow = 0;
                    while (xmlReader.Read())
                    {
                        if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.LocalName == XlsxWorksheet.NWorksheet)
                        {
                            // grab the namespaceuri from the worksheet element
                            _namespaceUri = xmlReader.NamespaceURI;
                        }

                        if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.LocalName == XlsxWorksheet.NDimension)
                        {
                            string dimValue = xmlReader.GetAttribute(XlsxWorksheet.ARef);

                            var dimension = new XlsxDimension(dimValue);
                            if (dimension.IsRange)
                            {
                                Dimension = dimension;
                                foundDimension = true;

                                break;
                            }
                        }

                        // removed: Do not use col to work out number of columns as this is really for defining formatting, so may not contain all columns
                        /*if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_col)
                            cols++;*/

                        if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.LocalName == XlsxWorksheet.NRow)
                        {
                            biggestColumn = Math.Max(biggestColumn, cellElementsInRow);
                            cellElementsInRow = 0;
                            rows++;
                        }

                        // check cells so we can find size of sheet if can't work it out from dimension or col elements (dimension should have been set before the cells if it was available)
                        // ditto for cols
                        if (cols == 0 && xmlReader.NodeType == XmlNodeType.Element && xmlReader.LocalName == XlsxWorksheet.NC)
                        {
                            cellElementsInRow++;

                            var refAttribute = xmlReader.GetAttribute(XlsxWorksheet.AR);

                            if (refAttribute != null)
                            {
                                int column;
                                ReferenceHelper.ParseReference(refAttribute, out column);
                                if (column > biggestColumn)
                                    biggestColumn = column;
                            }
                        }
                    }

                    biggestColumn = Math.Max(biggestColumn, cellElementsInRow);

                    // if we didn't get a dimension element then use the calculated rows/cols to create it
                    if (!foundDimension)
                    {
                        if (cols == 0)
                            cols = biggestColumn;

                        if (rows == 0 || cols == 0)
                        {
                            IsEmpty = true;
                            return;
                        }

                        Dimension = new XlsxDimension(rows, cols);
                    }
                }
            }

            using (var sheetStream = Document.GetWorksheetStream(Path))
            {
                // read up to the sheetData element. if this element is empty then there aren't any rows and we need to null out dimension
                using (var xmlReader = XmlReader.Create(sheetStream))
                {
                    xmlReader.ReadToFollowing(XlsxWorksheet.NSheetData, _namespaceUri);
                    if (xmlReader.IsEmptyElement)
                    {
                        IsEmpty = true;
                    }
                }
            }
        }

        private class XlsxRowBlock
        {
            public int RowIndex { get; set; }

            public object[] Values { get; set; }
        }
    }
}
