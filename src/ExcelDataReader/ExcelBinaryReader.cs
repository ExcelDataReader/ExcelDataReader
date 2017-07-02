using System.IO;
using ExcelDataReader.Core.BinaryFormat;

namespace ExcelDataReader
{
    /// <summary>
    /// ExcelDataReader Class
    /// </summary>
    internal class ExcelBinaryReader : ExcelDataReader<XlsWorkbook, XlsWorksheet>
    {
        public ExcelBinaryReader(Stream stream, ExcelReaderConfiguration configuration)
            : base(configuration)
        {
            Stream = stream;
            Workbook = new XlsWorkbook(stream, Configuration.FallbackEncoding);

            // By default, the data reader is positioned on the first result.
            Reset();
        }

        private Stream Stream { get; set; }

        public override void Close()
        {
            base.Close();

            Stream?.Dispose();
            Stream = null;
            Workbook = null;
        }
    }
}
