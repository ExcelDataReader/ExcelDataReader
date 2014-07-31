namespace ExcelDataReader.Silverlight.Data.Example
{
	using System.Collections.Generic;
	using Data;

	public class DataRowCollection : List<IDataRow>, IDataRowCollection
	{
		internal DataRowCollection() {}
	}
}