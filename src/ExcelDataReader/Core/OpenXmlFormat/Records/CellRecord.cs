namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class CellRecord : Record
    {
        public CellRecord(int columnIndex, int xfIndex, object value, CellError? error)
        {
            ColumnIndex = columnIndex;
            XfIndex = xfIndex;
            Value = value;
            Error = error;
        }

        public int ColumnIndex { get; }

        public int XfIndex { get; }

        public object Value { get; }

        public CellError? Error { get; }
    }
}
