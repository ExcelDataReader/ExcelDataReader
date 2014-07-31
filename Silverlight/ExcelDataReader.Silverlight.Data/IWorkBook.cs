namespace ExcelDataReader.Silverlight.Data
{
	public interface IWorkBook
	{
		IWorkSheetCollection WorkSheets { get; }
		IWorkSheet CreateWorkSheet();
	}
}