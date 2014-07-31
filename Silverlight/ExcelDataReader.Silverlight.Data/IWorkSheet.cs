namespace ExcelDataReader.Silverlight.Data
{
	public interface IWorkSheet
	{
		string Name { get; set; }
		IDataColumnCollection Columns { get; }
		IDataColumn CreateDataColumn();
		IDataRowCollection Rows { get; }
		IDataRow CreateDataRow();
	}
}