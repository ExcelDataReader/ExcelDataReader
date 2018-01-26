﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core.CsvFormat
{
    internal class CsvWorksheet : IWorksheet
    {
        public CsvWorksheet(Stream stream, Encoding fallbackEncoding, char[] autodetectSeparators)
        {
            Stream = stream;
            Stream.Seek(0, SeekOrigin.Begin);
            try
            {
                // Try as UTF-8 first, or use BOM if present
                CsvAnalyzer.Analyze(Stream, autodetectSeparators, Encoding.UTF8, out var fieldCount, out var separator, out var encoding, out var bomLength, out var rowCount);
                FieldCount = fieldCount;
                RowCount = rowCount;
                Encoding = encoding;
                Separator = separator;
                BomLength = bomLength;
            }
            catch (DecoderFallbackException)
            {
                // If cannot parse as UTF-8, try fallback encoding
                Stream.Seek(0, SeekOrigin.Begin);

                CsvAnalyzer.Analyze(Stream, autodetectSeparators, fallbackEncoding, out var fieldCount, out var separator, out var encoding, out var bomLength, out var rowCount);
                FieldCount = fieldCount;
                RowCount = rowCount;
                Encoding = encoding;
                Separator = separator;
                BomLength = bomLength;
            }
        }

        public string Name => string.Empty;

        public string CodeName => null;

        public string VisibleState => null;

        public HeaderFooter HeaderFooter => null;

        public CellRange[] MergeCells => null;

        public int FieldCount { get; }

        public int RowCount { get; }

        public Stream Stream { get; }

        public Encoding Encoding { get; }

        public char Separator { get; }

        private int BomLength { get; set; }

        public NumberFormatString GetNumberFormatString(int index)
        {
            return null;
        }

        public IEnumerable<Row> ReadRows()
        {
            var bufferSize = 1024;
            var buffer = new byte[bufferSize];
            var rowIndex = 0;
            var csv = new CsvParser(Separator, Encoding);
            var skipBomBytes = BomLength;

            Stream.Seek(0, SeekOrigin.Begin);
            while (Stream.Position < Stream.Length)
            {
                var bytesRead = Stream.Read(buffer, 0, bufferSize);
                csv.ParseBuffer(buffer, skipBomBytes, bytesRead - skipBomBytes, out var bufferRows);

                skipBomBytes = 0; // Only skip bom on first iteration

                foreach (var row in GetReaderRows(rowIndex, bufferRows))
                {
                    yield return row;
                }

                rowIndex += bufferRows.Count;
            }

            csv.Flush(out var flushRows);
            foreach (var row in GetReaderRows(rowIndex, flushRows))
            {
                yield return row;
            }
        }

        private IEnumerable<Row> GetReaderRows(int rowIndex, List<List<string>> rows)
        {
            foreach (var row in rows)
            {
                var columnIndex = 0;

                var cells = new List<Cell>();
                foreach (var item in row)
                {
                    cells.Add(new Cell()
                    {
                        ColumnIndex = columnIndex,
                        Value = item
                    });

                    columnIndex++;
                }

                yield return new Row()
                {
                    Height = 12.75, // 255 twips
                    Cells = cells,
                    RowIndex = rowIndex
                };

                rowIndex++;
            }
        }
    }
}
