using System.Globalization;

#nullable enable

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class SheetRecord : Record
    {
        public SheetRecord(string name, uint id, string? rid, string visibleState, string? path)
        {
            Name = name;
            Id = id;
            Rid = rid;
            VisibleState = string.IsNullOrEmpty(visibleState) ? "visible" : visibleState.ToLower(CultureInfo.InvariantCulture);
            Path = path;
        }

        public string Name { get; }

        public string VisibleState { get; }

        public uint Id { get; }

        public string? Rid { get; }

        public string? Path { get; }
    }
}
