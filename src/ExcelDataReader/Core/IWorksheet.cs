namespace ExcelDataReader.Core;

/// <summary>
/// Common worksheet metadata and push-reader creation across supported formats.
/// </summary>
internal interface IWorksheet
{
    string Name { get; }

    string? CodeName { get; }

    string? VisibleState { get; }

    HeaderFooter? HeaderFooter { get; }

    int FieldCount { get; }

    int RowCount { get; }

    CellRange? Dimension { get; }

    CellRange[] MergeCells { get; }

    List<Column> ColumnWidths { get; }

    double DefaultRowHeight { get; }

    IWorksheetReader OpenReader();
}