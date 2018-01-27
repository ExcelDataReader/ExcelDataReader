using System;
using System.IO;
using System.Text;

namespace ExcelDataReader.Core.CsvFormat
{
    internal static class CsvAnalyzer
    {
        /// <summary>
        /// Reads completely through a CSV stream to determine encoding, separator, field count and row count. 
        /// Uses fallbackEncoding if there is no BOM. Throws DecoderFallbackException if there are invalid characters in the stream.
        /// Returns the separator whose average field count is closest to its max field count.
        /// </summary>
        public static void Analyze(Stream stream, char[] separators, Encoding fallbackEncoding, out int fieldCount, out char autodetectSeparator, out Encoding autodetectEncoding, out int bomLength, out int rowCount)
        {
            var bufferSize = 1024;
            var probeSize = 16;
            var buffer = new byte[bufferSize];
            var bytesRead = stream.Read(buffer, 0, probeSize);

            autodetectEncoding = GetEncodingFromBom(buffer, out bomLength);
            if (autodetectEncoding == null)
            {
                autodetectEncoding = fallbackEncoding;
            }

            if (separators == null || separators.Length == 0)
            {
                separators = new char[] { '\0' };
            }

            var separatorInfos = new SeparatorInfo[separators.Length];
            for (var i = 0; i < separators.Length; i++)
            {
                separatorInfos[i] = new SeparatorInfo();
                separatorInfos[i].Buffer = new CsvParser(separators[i], autodetectEncoding);
            }

            ParseSeparatorsBuffer(buffer, bomLength, bytesRead - bomLength, separators, separatorInfos);

            while (stream.Position < stream.Length)
            {
                bytesRead = stream.Read(buffer, 0, bufferSize);
                ParseSeparatorsBuffer(buffer, 0, bytesRead, separators, separatorInfos);
            }

            FlushSeparatorsBuffers(separators, separatorInfos);

            SeparatorInfo bestSeparatorInfo = separatorInfos[0];
            char bestSeparator = separators[0];
            double bestDistance = double.MaxValue;

            for (var i = 0; i < separators.Length; i++)
            {
                var separator = separators[i];
                var separatorInfo = separatorInfos[i];

                // Row has one column if there are no separators, there must be at least one separator to count
                if (separatorInfo.RowCount == 0 || separatorInfo.MaxFieldCount <= 1)
                {
                    continue;
                }

                var average = separatorInfo.SumFieldCount / (double)separatorInfo.RowCount;
                var dist = separatorInfo.MaxFieldCount - average;

                if (dist < bestDistance)
                {
                    bestDistance = dist;
                    bestSeparator = separator;
                    bestSeparatorInfo = separatorInfo;
                }
            }

            autodetectSeparator = bestSeparator;
            fieldCount = bestSeparatorInfo.MaxFieldCount;
            rowCount = bestSeparatorInfo.RowCount;
        }

        private static void ParseSeparatorsBuffer(byte[] bytes, int offset, int count, char[] separators, SeparatorInfo[] separatorInfos)
        {
            for (var i = 0; i < separators.Length; i++)
            {
                var separator = separators[i];
                SeparatorInfo separatorInfo = separatorInfos[i];

                separatorInfo.Buffer.ParseBuffer(bytes, offset, count, out var rows);

                foreach (var row in rows)
                {
                    separatorInfo.MaxFieldCount = Math.Max(separatorInfo.MaxFieldCount, row.Count);
                    separatorInfo.SumFieldCount += row.Count;
                    separatorInfo.RowCount++;
                }
            }
        }

        private static void FlushSeparatorsBuffers(char[] separators, SeparatorInfo[] separatorInfos)
        {
            for (var i = 0; i < separators.Length; i++)
            {
                var separator = separators[i];
                SeparatorInfo separatorInfo = separatorInfos[i];

                separatorInfo.Buffer.Flush(out var rows);

                foreach (var row in rows)
                {
                    separatorInfo.MaxFieldCount = Math.Max(separatorInfo.MaxFieldCount, row.Count);
                    separatorInfo.SumFieldCount += row.Count;
                    separatorInfo.RowCount++;
                }
            }
        }

        private static Encoding GetEncodingFromBom(byte[] bom, out int bomLength)
        {
            var encodings = new Encoding[]
            {
                Encoding.Unicode, Encoding.BigEndianUnicode, Encoding.UTF8
            };

            foreach (var encoding in encodings)
            {
                if (IsEncodingPreamble(bom, encoding, out int length))
                {
                    bomLength = length;
                    return encoding;
                }
            }

            bomLength = 0;
            return null;
        }

        private static bool IsEncodingPreamble(byte[] bom, Encoding encoding, out int bomLength)
        {
            bomLength = 0;
            var preabmle = encoding.GetPreamble();
            if (preabmle.Length > bom.Length)
                return false;
            var i = 0;
            for (; i < preabmle.Length; i++)
            {
                if (preabmle[i] != bom[i])
                    return false;
            }

            bomLength = i;
            return true;
        }

        private class SeparatorInfo
        {
            public int MaxFieldCount { get; set; }

            public int SumFieldCount { get; set; }

            public int RowCount { get; set; }

            public CsvParser Buffer { get; set; }
        }
    }
}
