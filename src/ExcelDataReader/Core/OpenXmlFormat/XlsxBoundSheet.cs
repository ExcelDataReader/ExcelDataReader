namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxBoundSheet
    {
        public XlsxBoundSheet(string name, int id, string rid, string visibleState)
        {
            Name = name;
            Id = id;
            Rid = rid;
            VisibleState = string.IsNullOrEmpty(visibleState) ? "visible" : visibleState.ToLower();
        }

        public string Name { get; }

        public string VisibleState { get; }

        public int Id { get; }

        public string Rid { get; set; }

        public string Path { get; set; }
    }
}
