namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class SheetDimRecord : Record
    {
        public SheetDimRecord(CellRange range)
        {
            Range = range;
        }

        public CellRange Range { get; }
    }
}
