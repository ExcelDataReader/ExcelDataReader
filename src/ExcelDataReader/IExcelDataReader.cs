using System.Data;

namespace ExcelDataReader;

/// <summary>
/// The ExcelDataReader interface.
/// </summary>
public interface IExcelDataReader : IDataReader
{
    /// <summary>
    /// Gets the sheet name.
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Gets the sheet VBA code name.
    /// </summary>
    string CodeName { get; }

    /// <summary>
    /// Gets the sheet visible state.
    /// </summary>
    string VisibleState { get; }

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
    HeaderFooter HeaderFooter { get; }

    /// <summary>
    /// Gets the list of merged cell ranges.
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
    CellRange Dimension { get; }

    /// <summary>
    /// Gets the height of the current row in points.
    /// </summary>
    double RowHeight { get; }

    /// <summary>
    /// Seeks to the first result.
    /// </summary>
    void Reset();

    /// <summary>
    /// Gets the number format for the specified field -or- <see langword="null"/> if there is no value.
    /// Returns Excel's original locale-independent built-in format strings.
    /// </summary>
    /// <param name="i">The index of the field to find.</param>
    /// <returns>The number format string of the specified field.</returns>
    string GetNumberFormatString(int i);

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
    string GetNumberFormatString(int i, IFormatProvider provider);

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
    /// <returns>The cell style.</returns>
    CellStyle GetCellStyle(int i);

    /// <summary>
    /// Gets the cell error.
    /// </summary>
    /// <param name="i">The index of the column to find.</param>
    /// <returns>The cell error, or null if no error.</returns>
    CellError? GetCellError(int i);
}