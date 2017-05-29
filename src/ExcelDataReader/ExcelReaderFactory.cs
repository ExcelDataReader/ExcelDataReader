using System;
using System.IO;
using ExcelDataReader.Core.BinaryFormat;

namespace Excel
{
    /// <summary>
    /// The ExcelReader Factory
    /// </summary>
    public static class ExcelReaderFactory
    {
        /// <summary>
        /// Creates an instance of <see cref="ExcelBinaryReader"/> or <see cref="ExcelOpenXmlReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <param name="convertOADates">If <see langword="true"/> convert OA dates to <see cref="DateTime"/>. Only applicable to binary (xls) files.</param>
        /// <param name="readOption">The read option to use for binary (xls) files.</param>
        /// <returns>The excel data reader.</returns>
        public static IExcelDataReader CreateReader(Stream fileStream, bool convertOADates = true, ReadOption readOption = ReadOption.Strict)
        {
            XlsHeader header = new XlsHeader(fileStream);
            if (header.IsSignatureValid)
                return new ExcelBinaryReader(fileStream, header, convertOADates, readOption);

            if (header.IsRawBiffStream)
                throw new NotSupportedException("File appears to be a raw BIFF stream which isn't supported (BIFF" + header.RawBiffVersion + ").");

            return CreateOpenXmlReader(fileStream);
        }

        /// <summary>
        /// Creates an instance of <see cref="ExcelBinaryReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <returns>The excel data reader.</returns>
        public static IExcelDataReader CreateBinaryReader(Stream fileStream)
        {
            return CreateBinaryReader(fileStream, true, ReadOption.Strict);
        }

        /// <summary>
        /// Creates an instance of <see cref="ExcelBinaryReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <param name="option">The read option.</param>
        /// <returns>The excel data reader.</returns>
        public static IExcelDataReader CreateBinaryReader(Stream fileStream, ReadOption option)
        {
            return CreateBinaryReader(fileStream, true, option);
        }

        /// <summary>
        /// Creates an instance of <see cref="ExcelBinaryReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <param name="convertOADate">If true oa dates will be converer to <see cref="DateTime"/>.</param>
        /// <returns>The excel data reader.</returns>
        public static IExcelDataReader CreateBinaryReader(Stream fileStream, bool convertOADate)
        {
            return CreateBinaryReader(fileStream, convertOADate, ReadOption.Strict);
        }

        /// <summary>
        /// Creates an instance of <see cref="ExcelBinaryReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <param name="convertOADate">If true oa dates will be converer to <see cref="DateTime"/>.</param>
        /// <param name="readOption">The read option.</param>
        /// <returns>The excel data reader.</returns>
        public static IExcelDataReader CreateBinaryReader(Stream fileStream, bool convertOADate, ReadOption readOption)
        {
            return new ExcelBinaryReader(fileStream, convertOADate, readOption);
        }

        /// <summary>
        /// Creates an instance of <see cref="ExcelOpenXmlReader"/>
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <returns>The excel data reader.</returns>
        public static IExcelDataReader CreateOpenXmlReader(Stream fileStream)
        {
            return new ExcelOpenXmlReader(fileStream);
        }
    }
}
