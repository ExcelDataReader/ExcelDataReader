namespace ExcelDataReader.Silverlight.Data.Example
{
	using System.Collections.Generic;
	using Data;

	public class DataColumnCollection : List<IDataColumn>, IDataColumnCollection
	{
		internal DataColumnCollection() {}
	}
}