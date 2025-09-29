namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal class SheetDimRecord : Record
    {
        public SheetDimRecord(CellRange range)
        {
            Range = range;
        }

        public CellRange Range { get; }
    }
}
