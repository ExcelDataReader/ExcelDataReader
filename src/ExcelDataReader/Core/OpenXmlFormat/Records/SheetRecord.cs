using System.Globalization;

namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class SheetRecord(string? name, uint id, string? rid, string? visibleState, string? path) : Record
{
    public string Name { get; } = name ?? string.Empty;

    public string VisibleState { get; } = visibleState is { Length: > 0 } state ? state.ToLower(CultureInfo.InvariantCulture) : "visible";

    public uint Id { get; } = id;

    public string? Rid { get; } = rid;

    public string? Path { get; } = path;
}
