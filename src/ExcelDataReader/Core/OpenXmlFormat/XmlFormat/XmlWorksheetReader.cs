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
            if (!Reader.IsStartElement(NWorksheet, ProperNamespaces.NsSpreadsheetMl))
            {
                yield break;
            }

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(NSheetData, ProperNamespaces.NsSpreadsheetMl))
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
                        if (Reader.IsStartElement(NRow, ProperNamespaces.NsSpreadsheetMl))
                        {
                            if (int.TryParse(Reader.GetAttribute(AR), out int arValue))
                                rowIndex = arValue - 1; // The row attribute is 1-based
                            else
                                rowIndex++;

#pragma warning disable CA1806 // Do not ignore method results
                            int.TryParse(Reader.GetAttribute(AHidden), out int hidden);
                            int.TryParse(Reader.GetAttribute(ACustomHeight), out int customHeight);
#pragma warning restore CA1806 // Do not ignore method results

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
                                if (Reader.IsStartElement(NC, ProperNamespaces.NsSpreadsheetMl))
                                {
                                    var cell = ReadCell(nextColumnIndex, ProperNamespaces.NsSpreadsheetMl);
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
                else if (Reader.IsStartElement(NMergeCells, ProperNamespaces.NsSpreadsheetMl))
                {
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        continue;
                    }

                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(NMergeCell, ProperNamespaces.NsSpreadsheetMl))
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
                else if (Reader.IsStartElement(NHeaderFooter, ProperNamespaces.NsSpreadsheetMl))
                {
                    var result = ReadHeaderFooter(ProperNamespaces.NsSpreadsheetMl);
                    if (result != null)
                        yield return new HeaderFooterRecord(result);
                }
                else if (Reader.IsStartElement(NCols, ProperNamespaces.NsSpreadsheetMl))
                {
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        continue;
                    }

                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(NCol, ProperNamespaces.NsSpreadsheetMl))
                        {
                            var min = Reader.GetAttribute(AMin);
                            var max = Reader.GetAttribute(AMax);
                            var width = Reader.GetAttribute(AWidth);
                            var customWidth = Reader.GetAttribute(ACustomWidth);
                            var hidden = Reader.GetAttribute(AHidden);

                            var maxVal = int.Parse(max, CultureInfo.InvariantCulture);
                            var minVal = int.Parse(min, CultureInfo.InvariantCulture);
                            double.TryParse(width, NumberStyles.Float, CultureInfo.InvariantCulture, out double widthVal);

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
                else if (Reader.IsStartElement(NSheetProperties, ProperNamespaces.NsSpreadsheetMl))
                {
                    var codeName = Reader.GetAttribute("codeName");
                    yield return new SheetPrRecord(codeName);

                    Reader.Skip();
                }
                else if (Reader.IsStartElement(NSheetFormatProperties, ProperNamespaces.NsSpreadsheetMl))
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

        private HeaderFooter ReadHeaderFooter(string nsSpreadsheetMl)
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
                if (Reader.IsStartElement(NOddHeader, nsSpreadsheetMl))
                {
                    headerFooter.OddHeader = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NOddFooter, nsSpreadsheetMl))
                {
                    headerFooter.OddFooter = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NEvenHeader, nsSpreadsheetMl))
                {
                    headerFooter.EvenHeader = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NEvenFooter, nsSpreadsheetMl))
                {
                    headerFooter.EvenFooter = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NFirstHeader, nsSpreadsheetMl))
                {
                    headerFooter.FirstHeader = Reader.ReadElementContentAsString();
                }
                else if (Reader.IsStartElement(NFirstFooter, nsSpreadsheetMl))
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

        private CellRecord ReadCell(int nextColumnIndex, string nsSpreadsheetMl)
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
                if (Reader.IsStartElement(NV, nsSpreadsheetMl))
                {
                    string rawValue = Reader.ReadElementContentAsString();
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, out value, out error);
                }
                else if (Reader.IsStartElement(NIs, nsSpreadsheetMl))
                {
                    string rawValue = StringHelper.ReadStringItem(Reader, nsSpreadsheetMl);
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, out value, out error);
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }

            return new CellRecord(columnIndex, xfIndex, value, error);

            static void ConvertCellValue(string rawValue, string aT, out object value, out CellError? error)
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

            // 2.5.97.2 BErr
            static CellError? ConvertError(string e) => e switch
            {
                "#NULL!" => CellError.NULL,
                "#DIV/0!" => CellError.DIV0,
                "#VALUE!" => CellError.VALUE,
                "#REF!" => CellError.REF,
                "#NAME?" => CellError.NAME,
                "#NUM!" => CellError.NUM,
                "#N/A" => CellError.NA,
                "#GETTING_DATA" => CellError.GETTING_DATA,
                _ => null,
            };
        }
    }
}
