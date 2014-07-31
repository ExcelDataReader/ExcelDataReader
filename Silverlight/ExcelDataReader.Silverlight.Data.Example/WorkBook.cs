namespace ExcelDataReader.Silverlight.Data.Example
{
	using Data;

	public class WorkBook : IWorkBook
	{
		internal WorkBook()
		{
			WorkSheets = new WorkSheetCollection();
		}

		public IWorkSheetCollection WorkSheets { get; private set; }

		public IWorkSheet CreateWorkSheet()
		{
			return new WorkSheet();
		}

		public string DataSetName { get; set; }
	}
}
