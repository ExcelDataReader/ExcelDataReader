namespace ExcelDataReader.Silverlight.Data.Example
{
	using System.Collections;
	using Data;

	public class DataRow : IDataRow
	{
		internal DataRow() {}

		public IList Values { get; set; }
	}
}