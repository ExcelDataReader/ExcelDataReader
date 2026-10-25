namespace ExcelDataReader;

/// <summary>
/// Immutable descriptive properties stored in a workbook. Missing or unsupported values are null.
/// </summary>
public sealed class DocumentProperties
{
    internal DocumentProperties()
    {
    }

    /// <summary>Gets the document title.</summary>
    public string? Title { get; internal init; }

    /// <summary>Gets the document subject.</summary>
    public string? Subject { get; internal init; }

    /// <summary>Gets the document author.</summary>
    public string? Author { get; internal init; }

    /// <summary>Gets the keywords as stored, without splitting them.</summary>
    public string? Keywords { get; internal init; }

    /// <summary>Gets the document description (the Comments property in Excel).</summary>
    public string? Description { get; internal init; }

    /// <summary>Gets the name stored for the last person to save the document.</summary>
    public string? LastModifiedBy { get; internal init; }

    /// <summary>Gets the document creation time in UTC.</summary>
    public DateTime? Created { get; internal init; }

    /// <summary>Gets the document modification time in UTC.</summary>
    public DateTime? Modified { get; internal init; }

    /// <summary>Gets the last printing time in UTC.</summary>
    public DateTime? LastPrinted { get; internal init; }

    internal static DocumentProperties Empty { get; } = new();
}
