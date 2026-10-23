namespace ExcelDataReader.Core;

internal interface IWorksheetReader : IDisposable
{
    // Paused means the last callback was accepted and more consumption is needed;
    // EndOfInput means no callbacks remain. Handler mutations must survive the call,
    // including when a reader forwards callbacks through another struct handler.
    WorksheetReadStatus Consume<THandler>(ref THandler handler)
        where THandler : struct, IWorksheetHandler;
}