using System.Globalization;

#nullable enable

namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class SheetRecord(string name, uint id, string? rid, string visibleState, string? path) : Record
{
    public string Name { get; } = name;

    public string VisibleState { get; } = string.IsNullOrEmpty(visibleState) ? "visible" : visibleState.ToLower(CultureInfo.InvariantCulture);

    public uint Id { get; } = id;

    public string? Rid { get; } = rid;

    public string? Path { get; } = path;
}
