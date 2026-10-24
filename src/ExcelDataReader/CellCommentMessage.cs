namespace ExcelDataReader;

/// <summary>
/// An immutable plain-text message in a cell comment.
/// </summary>
public sealed class CellCommentMessage
{
    internal CellCommentMessage(string? author, string text)
    {
        Author = author;
        Text = text;
    }

    /// <summary>
    /// Gets the author's display name, or null when the format does not provide one.
    /// </summary>
    public string? Author { get; }

    /// <summary>
    /// Gets the comment text without rich-text formatting.
    /// </summary>
    public string Text { get; }
}