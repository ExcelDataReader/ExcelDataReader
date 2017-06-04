// ReSharper disable InconsistentNaming
namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxWorksheet
    {
        public const string NDimension = "dimension";
        public const string NWorksheet = "worksheet";
        public const string NRow = "row";
        public const string NCol = "col";
        public const string NC = "c"; // cell
        public const string NV = "v";
        public const string NT = "t";
        public const string ARef = "ref";
        public const string AR = "r";
        public const string AT = "t";
        public const string AS = "s";
        public const string NSheetData = "sheetData";
        public const string NInlineStr = "inlineStr";

        public XlsxWorksheet(string name, int id, string rid, string visibleState)
        {
            Name = name;
            Id = id;
            Rid = rid;
            VisibleState = string.IsNullOrEmpty(visibleState) ? "visible" : visibleState.ToLower();
        }

        public bool IsEmpty { get; set; }

        public XlsxDimension Dimension { get; set; }

        public int ColumnsCount => IsEmpty ? 0 : (Dimension?.LastCol ?? -1);

        public int RowsCount => Dimension == null ? -1 : Dimension.LastRow - Dimension.FirstRow + 1;

        public string Name { get; }

        public string VisibleState { get; }

        public int Id { get; }

        public string Rid { get; set; }

        public string Path { get; set; }
    }
}
