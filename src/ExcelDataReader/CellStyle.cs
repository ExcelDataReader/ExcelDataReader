namespace ExcelDataReader;

/// <summary>
/// Horizontal alignment. 
/// </summary>
public enum HorizontalAlignment
{
    /// <summary>
    /// General.
    /// </summary>
    General,

    /// <summary>
    /// Left.
    /// </summary>
    Left,

    /// <summary>
    /// Center.
    /// </summary>
    Center = 2,

    /// <summary>
    /// Right.
    /// </summary>
    Right,

    /// <summary>
    /// Filled.
    /// </summary>
    Filled,

    /// <summary>
    /// Justified.
    /// </summary>
    Justified,

    /// <summary>
    /// Centered across selection.
    /// </summary>
    CenteredAcrossSelection,

    /// <summary>
    /// Distributed.
    /// </summary>
    Distributed,

    /// <summary>
    /// Same as <see cref="Center"/>.
    /// </summary>
    /// <remarks>
    /// This is an alias for <see cref="Center"/> to maintain compatibility with older versions of the library.
    /// It is recommended to use <see cref="Center"/> for clarity in new code.
    /// </remarks>
    Centered = Center,
}

/// <summary>
/// Vertical alignment.
/// </summary>
public enum VerticalAlignment
{
    /// <summary>
    /// Top.
    /// </summary>
    Top,

    /// <summary>
    /// Center.
    /// </summary>
    Center,

    /// <summary>
    /// Bottom.
    /// </summary>
    Bottom,

    /// <summary>
    /// Justify.
    /// </summary>
    Justify,

    /// <summary>
    /// Distributed.
    /// </summary>
    Distributed,
}

/// <summary>
/// Holds style information for a cell.
/// </summary>
public class CellStyle
{
    /// <summary>
    /// Gets the font index.
    /// </summary>
    public int FontIndex { get; internal set; }

    /// <summary>
    /// Gets the number format index.
    /// </summary>
    public int NumberFormatIndex { get; internal set; }

    /// <summary>
    /// Gets the indent level.
    /// </summary>
    public int IndentLevel { get; internal set; }

    /// <summary>
    /// Gets the horizontal alignment.
    /// </summary>
    public HorizontalAlignment HorizontalAlignment { get; internal set; }

    /// <summary>
    /// Gets the vertical alignment.
    /// </summary>
    public VerticalAlignment VerticalAlignment { get; internal set; }

    /// <summary>
    /// Gets a value indicating whether the cell is hidden.
    /// </summary>
    public bool Hidden { get; internal set; }

    /// <summary>
    /// Gets a value indicating whether the cell is locked.
    /// </summary>
    public bool Locked { get; internal set; }
}