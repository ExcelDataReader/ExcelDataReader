using System.IO;

#nullable enable

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class SheetRecord : Record
    {
        public SheetRecord(string name, uint id, string? rid, string visibleState)
        {
            Name = name;
            Id = id;
            Rid = rid;
            VisibleState = string.IsNullOrEmpty(visibleState) ? "visible" : visibleState.ToLower();
        }

        public string Name { get; }

        public string VisibleState { get; }

        public uint Id { get; }

        public string? Rid { get; set; }

        public string? Path { get; set; }
    }
}
