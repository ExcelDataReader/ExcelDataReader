using System.Xml;

namespace ExcelDataReader.Core.OpenXmlFormat.Records;

/// <summary>
/// Represents a relationship record.
/// </summary>
internal sealed class RelRecord(string id, string type, string target) : Record
{
    public string Id { get; } = id;

    public string Type { get; } = type;

    public string Target { get; } = target;
}