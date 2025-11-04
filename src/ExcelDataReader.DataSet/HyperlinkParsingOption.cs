namespace ExcelDataReader;

/// <summary>
/// Describes how hyperlink cells should be parsed when reading Excel files.
/// 
/// </summary>
public enum HyperlinkParsingOption
{
    /// <summary>
    /// Only parse the display text of the hyperlink.
    /// </summary>
    DisplayText,

    /// <summary>
    /// Only parse the URL of the hyperlink.
    /// </summary>
    URL,

    /// <summary>
    /// Parse the hyperlink as a <see cref="Tuple{T1, T2}"/> 
    /// containing the display text (<see cref="object"/>) and the URL (<see cref="object"/>).
    /// </summary>
    Tuple
}
