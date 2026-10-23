namespace ExcelDataReader.Core;

internal interface IWorksheetReader : IDisposable
{
    WorksheetReadStatus Consume<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler;
}