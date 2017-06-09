using System;
using System.IO;
using ExcelDataReader.Core.BinaryFormat;

namespace ExcelDataReader
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
            var probe = new byte[8];
            fileStream.Read(probe, 0, probe.Length);
            fileStream.Seek(0, SeekOrigin.Begin);

            if (!XlsDocument.CheckRawBiffStream(probe, out int version))
            {
                version = -1;
            }

            // MUST be set to the value 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1.
            if (version != -1 || (probe[0] == 0xD0 && probe[1] == 0xCF)) { 
                return new ExcelBinaryReader(fileStream, convertOADates, readOption);
            }

            // zip files start with 'PK'
            if (probe[0] == 0x50 && probe[1] == 0x4B) { 
                return CreateOpenXmlReader(fileStream);
            }

            throw new NotSupportedException("Unknown file format");
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
