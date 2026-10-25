namespace ExcelDataReader.Core;

internal sealed class DocumentPropertiesBuilder
{
    public string? Title { get; set; }

    public string? Subject { get; set; }

    public string? Author { get; set; }

    public string? Keywords { get; set; }

    public string? Description { get; set; }

    public string? LastModifiedBy { get; set; }

    public DateTime? Created { get; set; }

    public DateTime? Modified { get; set; }

    public DateTime? LastPrinted { get; set; }

    public DocumentProperties Build() => new()
    {
        Title = Title,
        Subject = Subject,
        Author = Author,
        Keywords = Keywords,
        Description = Description,
        LastModifiedBy = LastModifiedBy,
        Created = Created,
        Modified = Modified,
        LastPrinted = LastPrinted,
    };
}
