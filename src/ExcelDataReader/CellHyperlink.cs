namespace ExcelDataReader;

/// <summary>
/// An immutable hyperlink, independent of the displayed cell value.
/// </summary>
public sealed class CellHyperlink
{
    internal CellHyperlink(string address, string? tooltip)
    {
        Address = address;
        Tooltip = tooltip;
    }

    /// <summary>
    /// Gets the complete destination, including any fragment or within-document location.
    /// Relative paths and escaping are preserved, and the destination is not validated or resolved.
    /// An empty address can accompany a tooltip-only hyperlink.
    /// </summary>
    public string Address { get; }

    /// <summary>
    /// Gets the stored tooltip, or null when none is present.
    /// </summary>
    public string? Tooltip { get; }
}
