namespace ExcelDataReader.Silverlight.Data.Example
{
	using Data;

	public class WorkBookFactory : IWorkBookFactory
	{
		public IWorkBook CreateWorkBook()
		{
			return new WorkBook();
		}
	}
}