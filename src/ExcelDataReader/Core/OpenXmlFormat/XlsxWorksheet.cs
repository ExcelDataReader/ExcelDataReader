using System;
using System.Collections.Generic;
using ExcelDataReader.Core.NumberFormat;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxWorksheet : IWorksheet
    {
        public XlsxWorksheet(ZipWorker document, XlsxWorkbook workbook, SheetRecord refSheet)
        {
            Document = document;
            Workbook = workbook;

            Name = refSheet.Name;
            Id = refSheet.Id;
            Rid = refSheet.Rid;
            VisibleState = refSheet.VisibleState;
            Path = refSheet.Path;
            DefaultRowHeight = 15;

            if (string.IsNullOrEmpty(Path))
                return;

            using var sheetStream = Document.GetWorksheetReader(Path);
            if (sheetStream == null)
                return;

            int rowIndexMaximum = int.MinValue;
            int columnIndexMaximum = int.MinValue;

            List<Column> columnWidths = new List<Column>();
            List<CellRange> cellRanges = new List<CellRange>();

            bool inSheetData = false;

            Record record;
            while ((record = sheetStream.Read()) != null)
            {
                switch (record)
                {
                    case SheetDataBeginRecord _:
                        inSheetData = true;
                        break;
                    case SheetDataEndRecord _:
                        inSheetData = false;
                        break;
                    case RowHeaderRecord row when inSheetData:
                        rowIndexMaximum = Math.Max(rowIndexMaximum, row.RowIndex);
                        break;
                    case CellRecord cell when inSheetData:
                        columnIndexMaximum = Math.Max(columnIndexMaximum, cell.ColumnIndex);
                        break;
                    case ColumnRecord column:
                        columnWidths.Add(column.Column);
                        break;
                    case SheetFormatPrRecord sheetFormatProperties:
                        if (sheetFormatProperties.DefaultRowHeight != null)
                            DefaultRowHeight = sheetFormatProperties.DefaultRowHeight.Value;
                        break;
                    case SheetPrRecord sheetProperties:
                        CodeName = sheetProperties.CodeName;
                        break;
                    case MergeCellRecord mergeCell:
                        cellRanges.Add(mergeCell.Range);
                        break;
                    case HeaderFooterRecord headerFooter:
                        HeaderFooter = headerFooter.HeaderFooter;
                        break;
                }
            }

            ColumnWidths = columnWidths.ToArray();
            MergeCells = cellRanges.ToArray();

            if (rowIndexMaximum != int.MinValue && columnIndexMaximum != int.MinValue)
            {
                FieldCount = columnIndexMaximum + 1;
                RowCount = rowIndexMaximum + 1;
            }
        }

        public int FieldCount { get; }

        public int RowCount { get; }

        public string Name { get; }

        public string CodeName { get; }

        public string VisibleState { get; }

        public HeaderFooter HeaderFooter { get; }

        public double DefaultRowHeight { get; }

        public uint Id { get; }

        public string Rid { get; set; }

        public string Path { get; set; }

        public CellRange[] MergeCells { get; }

        public Column[] ColumnWidths { get; }

        private ZipWorker Document { get; }

        private XlsxWorkbook Workbook { get; }

        public IEnumerable<Row> ReadRows()
        {
            if (string.IsNullOrEmpty(Path))
                yield break;

            using var sheetStream = Document.GetWorksheetReader(Path);
            if (sheetStream == null)
                yield break;

            var rowIndex = 0;
            List<Cell> cells = null;
            double height = 0;

            bool inSheetData = false;
            Record record;
            while ((record = sheetStream.Read()) != null)
            {
                switch (record)
                {
                    case SheetDataBeginRecord _:
                        inSheetData = true;
                        break;
                    case SheetDataEndRecord _:
                        inSheetData = false;
                        break;
                    case RowHeaderRecord row when inSheetData:
                        int currentRowIndex = row.RowIndex;

                        if (cells != null && rowIndex != currentRowIndex)
                        {
                            yield return new Row(rowIndex++, height, cells);
                            cells = null;
                        }

                        if (cells == null)
                        {
                            height = row.Hidden ? 0 : row.Height ?? DefaultRowHeight;
                            cells = new List<Cell>();
                        }

                        for (; rowIndex < currentRowIndex; rowIndex++)
                        {
                            yield return new Row(rowIndex, DefaultRowHeight, new List<Cell>());
                        }

                        break;
                    case CellRecord cell when inSheetData:
                        // TODO What if we get a cell without a row?
                        var extendedFormat = Workbook.GetEffectiveCellStyle(cell.XfIndex, 0);
                        cells.Add(new Cell(cell.ColumnIndex, ConvertCellValue(cell.Value, extendedFormat.NumberFormatIndex), extendedFormat, cell.Error));
                        break;
                }
            }

            if (cells != null)
                yield return new Row(rowIndex, height, cells);
        }

        private object ConvertCellValue(object value, int numberFormatIndex)
        {
            switch (value)
            {
                case int sstIndex:
                    if (sstIndex >= 0 && sstIndex < Workbook.SST.Count)
                    {
                        return Helpers.ConvertEscapeChars(Workbook.SST[sstIndex]);
                    }

                    return null;

                case double number:
                    var format = Workbook.GetNumberFormatString(numberFormatIndex);
                    if (format != null)
                    {
                        if (format.IsDateTimeFormat)
                            return Helpers.ConvertFromOATime(number, Workbook.IsDate1904);
                        if (format.IsTimeSpanFormat)
                            return TimeSpan.FromDays(number);
                    }

                    return number;
                default:
                    return value;
            }
        }
    }
}
