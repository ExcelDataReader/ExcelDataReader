namespace ExcelDataReader;

/// <summary>
/// An immutable cell comment. A legacy note contains one message.
/// </summary>
public sealed class CellCommentThread
{
    internal CellCommentThread(string? author, string text)
    {
        Messages = Array.AsReadOnly(new[] { new CellCommentMessage(author, text) });
    }

    /// <summary>
    /// Gets the messages in this comment.
    /// </summary>
    public IReadOnlyList<CellCommentMessage> Messages { get; }
}