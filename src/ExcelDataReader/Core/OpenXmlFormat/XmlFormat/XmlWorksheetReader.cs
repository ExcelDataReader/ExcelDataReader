using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal sealed class XmlWorksheetReader : XmlRecordReader
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private const string NWorksheet = "worksheet";
        private const string NSheetData = "sheetData";
        private const string NRow = "row";
        private const string ARef = "ref";
        private const string AR = "r";
        private const string NV = "v";
        private const string NIs = "is";
        private const string AT = "t";
        private const string AS = "s";

        private const string NC = "c"; // cell
        private const string NInlineStr = "inlineStr";
        private const string NStr = "str";

        private const string NMergeCells = "mergeCells";

        private const string NSheetProperties = "sheetPr";
        private const string NSheetFormatProperties = "sheetFormatPr";
        private const string ADefaultRowHeight = "defaultRowHeight";

        private const string NHeaderFooter = "headerFooter";
        private const string ADifferentFirst = "differentFirst";
        private const string ADifferentOddEven = "differentOddEven";
        private const string NFirstHeader = "firstHeader";
        private const string NFirstFooter = "firstFooter";
        private const string NOddHeader = "oddHeader";
        private const string NOddFooter = "oddFooter";
        private const string NEvenHeader = "evenHeader";
        private const string NEvenFooter = "evenFooter";

        private const string NCols = "cols";
        private const string NCol = "col";
        private const string AMin = "min";
        private const string AMax = "max";
        private const string AHidden = "hidden";
        private const string AWidth = "width";
        private const string ACustomWidth = "customWidth";

        private const string NMergeCell = "mergeCell";

        private const string ACustomHeight = "customHeight";
        private const string AHt = "ht";

        public XmlWorksheetReader(XmlReader reader) 
            : base(reader)
        {
        }

        protected override IEnumerable<Record> ReadOverride()
        {
            if (!Reader.IsStartElement(NWorksheet, NsSpreadsheetMl))
            {
                yield break;
            }

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(NSheetData, NsSpreadsheetMl))
                {
                    yield return new SheetDataBeginRecord();
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        yield return new SheetDataEndRecord();
                        continue;
                    }

                    int rowIndex = -1;
                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(NRow, NsSpreadsheetMl))
                        {
                            if (int.TryParse(Reader.GetAttribute(AR), out int arValue))
                                rowIndex = arValue - 1; // The row attribute is 1-based
                            else
                                rowIndex++;

                            int.TryParse(Reader.GetAttribute(AHidden), out int hidden);
                            int.TryParse(Reader.GetAttribute(ACustomHeight), out int customHeight);
                            double? height;
                            if (customHeight != 0 && double.TryParse(Reader.GetAttribute(AHt), NumberStyles.Any, CultureInfo.InvariantCulture, out var ahtValue))
                                height = ahtValue;
                            else
                                height = null;

                            yield return new RowHeaderRecord(rowIndex, hidden != 0, height);

                            if (!XmlReaderHelper.ReadFirstContent(Reader))
                            {
                                continue;
                            }

                            int nextColumnIndex = 0;
                            while (!Reader.EOF)
                            {
                                if (Reader.IsStartElement(NC, NsSpreadsheetMl))
                                {
                                    var cell = ReadCell(nextColumnIndex);
                                    nextColumnIndex = cell.ColumnIndex + 1;
                                    yield return cell;
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

                    yield return new SheetDataEndRecord();
                }
                else if (Reader.IsStartElement(NMergeCells, NsSpreadsheetMl))
                {
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        continue;
                    }

                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(NMergeCell, NsSpreadsheetMl))
                        {
                            var cellRefs = Reader.GetAttribute(ARef);
                            yield return new MergeCellRecord(new CellRange(cellRefs));

                            Reader.Skip();
                        }
                        else if (!XmlReaderHelper.SkipContent(Reader))
                        {
                            break;
                        }
                    }
                }
                else if (Reader.IsStartElement(NHeaderFooter, NsSpreadsheetMl))
                {
                    var result = ReadHeaderFooter();
                    if (result != null)
                        yield return new HeaderFooterRecord(result);
                }
                else if (Reader.IsStartElement(NCols, NsSpreadsheetMl))
                {
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        continue;
                    }

                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(NCol, NsSpreadsheetMl))
                        {
                            var min = Reader.GetAttribute(AMin);
                            var max = Reader.GetAttribute(AMax);
                            var width = Reader.GetAttribute(AWidth);
                            var customWidth = Reader.GetAttribute(ACustomWidth);
                            var hidden = Reader.GetAttribute(AHidden);

                            var maxVal = int.Parse(max);
                            var minVal = int.Parse(min);
                            var widthVal = double.Parse(width, CultureInfo.InvariantCulture);

                            // Note: column indexes need to be converted to be zero-indexed
                            yield return new ColumnRecord(new Column(minVal - 1, maxVal - 1, hidden == "1", customWidth == "1" ? (double?)widthVal : null));

                            Reader.Skip();
                        }
                        else if (!XmlReaderHelper.SkipContent(Reader))
                        {
                            break;
                        }
                    }
                }
                else if (Reader.IsStartElement(NSheetProperties, NsSpreadsheetMl))
                {
                    var codeName = Reader.GetAttribute("codeName");
                    yield return new SheetPrRecord(codeName);

                    Reader.Skip();
                }
                else if (Reader.IsStartElement(NSheetFormatProperties, NsSpreadsheetMl))
                {
                    if (double.TryParse(Reader.GetAttribute(ADefaultRowHeight), NumberStyles.Any, CultureInfo.InvariantCulture, out var defaultRowHeight))
                        yield return new SheetFormatPrRecord(defaultRowHeight);

                    Reader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }
        }

        private HeaderFooter ReadHeaderFooter()
        {
            var differentFirst = Reader.GetAttribute(ADifferentFirst) == "1";
            var differentOddEven = Reader.GetAttribute(ADifferentOddEven) == "1";

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                return null;
            }

            var headerFooter = new HeaderFooter(differentFirst, differentOddEven);

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(NOddHeader, NsSpreadsheetMl))
                {
                    headerFooter.OddHeader = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NOddFooter, NsSpreadsheetMl))
                {
                    headerFooter.OddFooter = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NEvenHeader, NsSpreadsheetMl))
                {
                    headerFooter.EvenHeader = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NEvenFooter, NsSpreadsheetMl))
                {
                    headerFooter.EvenFooter = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NFirstHeader, NsSpreadsheetMl))
                {
                    headerFooter.FirstHeader = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NFirstFooter, NsSpreadsheetMl))
                {
                    headerFooter.FirstFooter = Reader.ReadElementContentAsString();
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }

            return headerFooter;
        }

        private CellRecord ReadCell(int nextColumnIndex)
        {
            int columnIndex;
            int xfIndex = -1;

            var aS = Reader.GetAttribute(AS);
            var aT = Reader.GetAttribute(AT);
            var aR = Reader.GetAttribute(AR);

            if (ReferenceHelper.ParseReference(aR, out int referenceColumn, out _))
                columnIndex = referenceColumn - 1; // ParseReference is 1-based
            else
                columnIndex = nextColumnIndex;

            if (aS != null)
            {
                if (int.TryParse(aS, NumberStyles.Any, CultureInfo.InvariantCulture, out var styleIndex))
                {
                    xfIndex = styleIndex;
                }
            }

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                return new CellRecord(columnIndex, xfIndex, null, null);
            }

            object value = null;
            CellError? error = null;
            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(NV, NsSpreadsheetMl))
                {
                    string rawValue = Reader.ReadElementContentAsString();
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, out value, out error);
                }
                else if (Reader.IsStartElement(NIs, NsSpreadsheetMl))
                {
                    string rawValue = StringHelper.ReadStringItem(Reader);
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, out value, out error);
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }

            return new CellRecord(columnIndex, xfIndex, value, error);
        }

        private void ConvertCellValue(string rawValue, string aT, out object value, out CellError? error)
        {
            const NumberStyles style = NumberStyles.Any;
            var invariantCulture = CultureInfo.InvariantCulture;

            error = null;
            switch (aT)
            {
                case AS: //// if string
                    if (int.TryParse(rawValue, style, invariantCulture, out var sstIndex))
                    {
                        // TODO: Can we get here when the sstIndex is not a valid index in the SST list?
                        value = sstIndex;
                        return;
                    }

                    value = rawValue;
                    return;
                case NInlineStr: //// if string inline
                case NStr: //// if cached formula string
                    value = Helpers.ConvertEscapeChars(rawValue);
                    return;
                case "b": //// boolean
                    value = rawValue == "1";
                    return;
                case "d": //// ISO 8601 date
                    if (DateTime.TryParseExact(rawValue, "yyyy-MM-dd", invariantCulture, DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite, out var date))
                    {
                        value = date;
                        return;
                    }

                    value = rawValue;
                    return;
                case "e": //// error
                    error = ConvertError(rawValue);
                    value = null;
                    return;
                default:
                    if (double.TryParse(rawValue, style, invariantCulture, out double number))
                    {
                        value = number;
                        return;
                    }

                    value = rawValue;
                    return;
            }
        }

        private CellError? ConvertError(string e)
        {
            // 2.5.97.2 BErr
            switch (e)
            {
                case "#NULL!":
                    return CellError.NULL;
                case "#DIV/0!":
                    return CellError.DIV0;
                case "#VALUE!":
                    return CellError.VALUE;
                case "#REF!":
                    return CellError.REF;
                case "#NAME?":
                    return CellError.NAME;
                case "#NUM!":
                    return CellError.NUM;
                case "#N/A":
                    return CellError.NA;
                case "#GETTING_DATA":
                    return CellError.GETTING_DATA;
                default:
                    return null;
            }
        }
    }
}
