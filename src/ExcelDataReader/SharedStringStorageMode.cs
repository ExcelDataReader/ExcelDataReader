namespace ExcelDataReader;

/// <summary>
/// Specifies how shared strings in XLSX, XLSB, and BIFF8 XLS workbooks are stored.
/// </summary>
public enum SharedStringStorageMode
{
    /// <summary>Keep the entire shared string table in memory without temporary SST files.</summary>
    Default,

    /// <summary>Use the normal string table until its accounted capacity exceeds the SST budget, then spill to temporary files.</summary>
    SpillToDisk,
}
