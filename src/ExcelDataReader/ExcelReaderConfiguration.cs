﻿using System.Text;

namespace ExcelDataReader
{
    /// <summary>
    /// Configuration options for an instance of ExcelDataReader.
    /// </summary>
    public class ExcelReaderConfiguration
    {
        /// <summary>
        /// Gets or sets a value indicating the encoding to use when the input XLS lacks a CodePage record, 
        /// or when the input CSV lacks a BOM and does not parse as UTF8. Default: cp1252. (XLS BIFF2-5 and CSV only)
        /// </summary>
        public Encoding FallbackEncoding { get; set; } = Encoding.GetEncoding(1252);

        /// <summary>
        /// Gets or sets the password used to open password protected workbooks.
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// Gets or sets an array of CSV separator candidates. The reader autodetects which best fits the input data. Default: , ; TAB | # (CSV only)
        /// </summary>
        public char[] AutodetectSeparators { get; set; } = new char[] { ',', ';', '\t', '|', '#' };

        /// <summary>
        /// Gets or sets a value indicating whether to leave the stream open after the IExcelDataReader object is disposed. Default: false
        /// </summary>
        public bool LeaveOpen { get; set; }

        /// <summary>
        /// Gets or sets a value indicating the number of row to analyze within a csv file. 
        /// 0 -> will analyze the enter file.
        /// > 0 will analyze at a minimum the number of row specified and will not return row count, it will be set to -1
        /// </summary>
        public int AnalyzeInitialCsvRows { get; set; }
    }
}
