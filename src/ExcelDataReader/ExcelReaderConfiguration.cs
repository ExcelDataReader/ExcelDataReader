using System.Text;

namespace ExcelDataReader;

/// <summary>
/// Configuration options for an instance of ExcelDataReader.
/// </summary>
public class ExcelReaderConfiguration
{
    /// <summary>
    /// Gets or sets a value indicating the encoding to use when the input XLS lacks a CodePage record,
    /// or when the input CSV lacks a BOM and does not parse as UTF8. Default: cp1252. (XLS BIFF2-5 and CSV only).
    /// </summary>
    public Encoding FallbackEncoding { get; set; } = Encoding.GetEncoding(1252);

    /// <summary>
    /// Gets or sets the password used to open password protected workbooks.
    /// </summary>
    public string? Password { get; set; }

    /// <summary>
    /// Gets or sets an array of CSV separator candidates. The reader autodetects which best fits the input data. Default: , ; TAB | # (CSV only).
    /// </summary>
    public char[] AutodetectSeparators { get; set; } = [',', ';', '\t', '|', '#'];

    /// <summary>
    /// Gets or sets a QuoteCharacter for CSV (Default '"').
    /// </summary>
    public char? QuoteChar { get; set; } = '"';

    /// <summary>
    /// Gets or sets a value indicating whether to trim white space values for CSV (Default 'true').
    /// </summary>
    public bool TrimWhiteSpace { get; set; } = true;

    /// <summary>
    /// Gets or sets a value indicating whether to leave the stream open after the IExcelDataReader object is disposed. Default: false.
    /// </summary>
    public bool LeaveOpen { get; set; }

    /// <summary>
    /// Gets or sets a value indicating the number of rows to analyze for encoding, separator and field count in a CSV.
    /// When set, this option causes the IExcelDataReader.RowCount property to throw an exception.
    /// Default: 0 - analyzes the entire file (CSV only, has no effect on other formats).
    /// </summary>
    public int AnalyzeInitialCsvRows { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether to skip the initial full-scan pass used to determine FieldCount and RowCount.
    /// When true, RowCount throws InvalidOperationException and FieldCount reflects the maximum column index seen so far,
    /// growing dynamically as rows are read. Default: false (XLS, XLSX/XLSB, and SpreadsheetML only, has no effect on CSV).
    /// </summary>
    public bool SinglePassMode { get; set; }

    /// <summary>
    /// Gets or sets shared string storage for XLSX, XLSB, and BIFF8 XLS. Default: Default.
    /// SpillToDisk retains the normal table until its accounted capacity exceeds SharedStringSpillThreshold,
    /// then migrates to temporary disk storage with a small bounded decoded cache.
    /// </summary>
    public SharedStringStorageMode SharedStringStorageMode { get; set; }

    /// <summary>
    /// Gets or sets the accounted SST storage threshold in bytes for SpillToDisk mode. Default: 64 MiB.
    /// Must be at least 1 MiB. Includes table capacity and strings, plus disk buffers and lookup
    /// cache after spill. Does not include parser temporaries, returned strings, rows, input
    /// buffering, or other workbook data.
    /// This is not a total managed-heap or process-memory limit.
    /// </summary>
    public long SharedStringSpillThreshold { get; set; } = 64L * 1024 * 1024;

    /// <summary>
    /// Gets or sets the temporary directory used when SpillToDisk storage spills. Default: the system
    /// temporary directory. Files are created only on spill and deleted when the reader closes.
    /// Temporary strings are plaintext, including for password-protected workbooks.
    /// Disk errors propagate without falling back to unbounded memory storage.
    /// </summary>
    public string? SharedStringTemporaryDirectory { get; set; }

    /// <summary>
    /// Gets or sets an escape character for CSV quoted fields (Default null - disabled).
    /// When set, this character inside a quoted field escapes the next character (e.g. set to '\' to support \" as an escaped quote).
    /// This is not part of RFC 4180; only enable it when reading files that use backslash-style escaping.
    /// </summary>
    public char? EscapeChar { get; set; }
}