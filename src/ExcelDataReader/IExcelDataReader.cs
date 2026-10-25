using System.Data;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader;

/// <summary>
/// The ExcelDataReader interface.
/// </summary>
public interface IExcelDataReader : IDataReader
{
    /// <summary>
    /// Gets the workbook's descriptive properties, loaded when the reader is opened.
    /// Missing or unsupported properties are null. CSV and raw BIFF have no document properties.
    /// </summary>
    DocumentProperties DocumentProperties { get; }

    /// <summary>
    /// Gets the sheet name.
    /// </summary>
    string? Name { get; }

    /// <summary>
    /// Gets the sheet VBA code name.
    /// </summary>
    string? CodeName { get; }

    /// <summary>
    /// Gets the sheet visible state.
    /// </summary>
    string? VisibleState { get; }

    /// <summary>
    /// Gets the selected CSV separator, or null for non-CSV readers, when separator detection
    /// is disabled, or when no result is selected.
    /// </summary>
    /// <remarks>
    /// Available immediately after creating a CSV reader, including after reaching the end of its rows.
    /// May contain multiple characters. The selected candidate need not occur in the input.
    /// </remarks>
    string? CsvSeparator { get; }

    /// <summary>
    /// Gets the active sheet.
    /// </summary>
    int ActiveSheet { get; }

    /// <summary>
    /// Gets a value indicating whether the worksheet is active.
    /// </summary>       
    bool IsActiveSheet { get; }

    /// <summary>
    /// Gets the sheet header and footer -or- <see langword="null"/> if none set.
    /// </summary>
    HeaderFooter? HeaderFooter { get; }

    /// <summary>
    /// Gets the list of merged cell ranges, or an empty array if there are none.
    /// </summary>
    CellRange[] MergeCells { get; }

    /// <summary>
    /// Gets the number of results (workbooks).
    /// </summary>
    int ResultsCount { get; }

    /// <summary>
    /// Gets the number of rows in the current result.
    /// </summary>
    int RowCount { get; }

    /// <summary>
    /// Gets the dimension of the current result.
    /// </summary>
    /// <remarks>Main use case is analysis of clipboard data. Potentially unreliable in other use cases.</remarks>
    CellRange? Dimension { get; }

    /// <summary>
    /// Gets the height of the current row in points.
    /// </summary>
    double RowHeight { get; }

    /// <summary>
    /// Seeks to the first result.
    /// </summary>
    void Reset();

    /// <summary>
    /// Restarts reading the current worksheet without changing the selected result.
    /// The next call to <see cref="IDataReader.Read"/> reads its first row.
    /// </summary>
    /// <exception cref="InvalidOperationException">The reader is closed or no worksheet is selected.</exception>
    /// <exception cref="ExcelReaderException">A recognized worksheet structure is invalid.</exception>
    void ResetWorksheet();

    /// <summary>
    /// Gets the number format for the specified field -or- <see langword="null"/> if there is no value.
    /// Returns Excel's original locale-independent built-in format strings.
    /// </summary>
    /// <param name="i">The index of the field to find.</param>
    /// <returns>The number format string of the specified field.</returns>
    string? GetNumberFormatString(int i);

    /// <summary>
    /// Gets the number format for the specified field using locale-dependent format patterns,
    /// -or- <see langword="null"/> if there is no value.
    /// </summary>
    /// <param name="i">The index of the field to find.</param>
    /// <param name="provider">
    /// An <see cref="IFormatProvider"/> (typically a <see cref="System.Globalization.CultureInfo"/>
    /// or <see cref="System.Globalization.DateTimeFormatInfo"/>) used to resolve locale-dependent
    /// built-in number format indices (14–17 and 20–22) to culture-specific date/time format strings.
    /// Pass <see langword="null"/> to use Excel's original locale-independent built-in format strings,
    /// equivalent to calling <see cref="GetNumberFormatString(int)"/>.
    /// </param>
    /// <returns>The number format string of the specified field.</returns>
    string? GetNumberFormatString(int i, IFormatProvider? provider);

    /// <summary>
    /// Gets the number format index for the specified field -or- -1 if there is no value.
    /// </summary>
    /// <param name="i">The index of the field to find.</param>
    /// <returns>The number format index of the specified field.</returns>
    int GetNumberFormatIndex(int i);

    /// <summary>
    /// Gets the width the specified column.
    /// </summary>
    /// <param name="i">The index of the column to find.</param>
    /// <returns>The width of the specified column.</returns>
    double GetColumnWidth(int i);

    /// <summary>
    /// Gets the cell style.
    /// </summary>
    /// <param name="i">The index of the column to find.</param>
    /// <returns>The immutable cell style, which may be shared by multiple cells.</returns>
    CellStyle GetCellStyle(int i);

    /// <summary>
    /// Gets the cell error.
    /// </summary>
    /// <param name="i">The index of the column to find.</param>
    /// <returns>The cell error, or null if no error.</returns>
    CellError? GetCellError(int i);

    /// <summary>
    /// Gets the legacy comment for the specified column on the current row, or null if none is present.
    /// Modern threaded comments are not currently read. CSV always returns null.
    /// </summary>
    /// <remarks>
    /// Loaded on first access, which may require an additional worksheet scan.
    /// Notes do not extend the reader's row or field counts.
    /// </remarks>
    /// <param name="i">The zero-based column index.</param>
    /// <returns>The immutable comment, or null.</returns>
    /// <exception cref="InvalidOperationException">There is no current row.</exception>
    /// <exception cref="IndexOutOfRangeException">The column is outside the current field count.</exception>
    /// <exception cref="ExcelReaderException">Supported comment data or its references are invalid.</exception>
    CellCommentThread? GetComment(int i);

    /// <summary>
    /// Gets the explicit stored hyperlink for the specified column on the current row,
    /// or null when there is no supported hyperlink. CSV always returns null.
    /// </summary>
    /// <remarks>
    /// Loaded on first access, which may require an additional worksheet scan.
    /// Hyperlinks do not change cell values or extend row or field counts. Formula-generated
    /// links are not read. Unsupported hyperlink representations are skipped.
    /// </remarks>
    /// <param name="i">The zero-based column index.</param>
    /// <returns>The immutable hyperlink, including tooltip-only links with an empty address, or null.</returns>
    /// <exception cref="InvalidOperationException">There is no current row.</exception>
    /// <exception cref="IndexOutOfRangeException">The column is outside the current field count.</exception>
    CellHyperlink? GetHyperlink(int i);
}