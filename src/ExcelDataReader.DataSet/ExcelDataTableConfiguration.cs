namespace ExcelDataReader;

/// <summary>
/// Processing configuration options and callbacks for AsDataTable().
/// </summary>
public class ExcelDataTableConfiguration
{
    /// <summary>
    /// Gets or sets a value indicating the prefix of generated column names.
    /// </summary>
    public string EmptyColumnNamePrefix { get; set; } = "Column";

    /// <summary>
    /// Gets or sets a value indicating whether to use a row from the data as column names.
    /// </summary>
    public bool UseHeaderRow { get; set; }

    /// <summary>
    /// Gets or sets a callback to determine which row is the header row. Only called when UseHeaderRow = true.
    /// </summary>
    public Action<IExcelDataReader> ReadHeaderRow { get; set; }

    /// <summary>
    /// Gets or sets a callback to allow a custom implementation of header reading.
    /// The returned dictionary will be used to construct the resulting DataTable.
    /// Each element of the dictionary specifies an index and column name pair.
    /// An example use of this would be to combine multiple header rows.
    /// NOTE: If this field is set, UseHeaderRow, EmptyColumnNamePrefix, and FilterColumn are ignored.
    /// </summary>
    public Func<IExcelDataReader, IReadOnlyDictionary<int, string>> ReadHeader { get; set; }

    /// <summary>
    /// Gets or sets a callback to determine whether to include the current row in the DataTable.
    /// </summary>
    public Func<IExcelDataReader, bool> FilterRow { get; set; }

    /// <summary>
    /// Gets or sets a callback to determine whether to include the specific column in the DataTable. Called once per column after reading the headers.
    /// </summary>
    public Func<IExcelDataReader, int, bool> FilterColumn { get; set; }

    /// <summary>
    /// Gets or sets a callback to determine whether to transform the cell value.
    /// </summary>
    public Func<IExcelDataReader, int, object, object> TransformValue { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether to parse the URL for hyperlink cells.
    /// </summary>
    public bool OverrideValueWithHyperlinkURL { get; set; }
}
