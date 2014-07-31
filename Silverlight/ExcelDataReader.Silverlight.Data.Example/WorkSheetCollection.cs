namespace ExcelDataReader.Silverlight.Data.Example
{
	using System.Collections.Generic;
	using Data;

	public class WorkSheetCollection : List<IWorkSheet>,  IWorkSheetCollection
	{
		internal WorkSheetCollection() {}
	}
}