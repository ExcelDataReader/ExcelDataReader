namespace ExcelDataReader.Silverlight.Data.Example
{
	using Data;

	public class WorkSheet : IWorkSheet
	{
		private readonly DataColumnCollection _Columns;
		private readonly DataRowCollection _Rows;

		internal WorkSheet()
		{
			_Columns = new DataColumnCollection();
			_Rows = new DataRowCollection();
		}

		public string Name { get; set; }

		public IDataColumnCollection Columns
		{
			get { return _Columns; }
		}

		public IDataColumn CreateDataColumn()
		{
			return new DataColumn();
		}

		public IDataRowCollection Rows
		{
			get { return _Rows; }
		}

		public IDataRow CreateDataRow()
		{
			return new DataRow();
		}
	}
}