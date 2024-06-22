namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class MergeCellRecord : Record
    {
        public MergeCellRecord(CellRange range) 
        {
            Range = range;
        }

        public CellRange Range { get; }
    }
}
