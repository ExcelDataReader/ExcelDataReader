using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat;

internal sealed class XlsxWorkbook : CommonWorkbook, IWorkbook<XlsxWorksheet>, IWorkbookPartSink, IStylesPartSink
{
    private readonly ZipWorker _zipWorker;
    private readonly ISharedStringStore _stringStore;

    public XlsxWorkbook(ZipWorker zipWorker, ExcelReaderConfiguration configuration)
    {
        _zipWorker = zipWorker;
        _stringStore = configuration.SharedStringStorageMode == SharedStringStorageMode.Default
            ? new SharedStringTable()
            : new AutoSharedStringStore(configuration);
        try
        {
            ReadWorkbook();
            ReadSharedStrings();
            ReadStyles();
        }
        catch (Exception exception)
        {
            ResourceCleanup.DisposeAll(exception, this);
            throw;
        }
    }

    public int SharedStringCount => _stringStore.Count;

    public bool IsDate1904 { get; private set; }

    public int ResultsCount => Sheets?.Count ?? -1;

    public int ActiveSheet { get; private set; }

    private List<SheetRecord> Sheets { get; } = [];

    public string GetSharedString(int index) => _stringStore.GetString(index);

    public IEnumerable<XlsxWorksheet> ReadWorksheets()
    {
        foreach (var sheet in Sheets)
            yield return new XlsxWorksheet(_zipWorker, this, sheet, SinglePassMode);
    }

    public void Dispose()
    {
        ResourceCleanup.DisposeAll(null, _stringStore, _zipWorker);
    }

    void IWorkbookPartSink.SetDate1904(bool date1904) => IsDate1904 = date1904;

    void IWorkbookPartSink.AddSheet(SheetRecord sheet) => Sheets.Add(sheet);

    void IWorkbookPartSink.SetActiveSheet(int activeSheet) => ActiveSheet = activeSheet;

    void IStylesPartSink.AddExtendedFormat(ExtendedFormat format) => ExtendedFormats.Add(format);

    void IStylesPartSink.AddCellStyleExtendedFormat(ExtendedFormat format) => CellStyleExtendedFormats.Add(format);

    void IStylesPartSink.AddNumberFormat(int formatIndexInFile, string formatString) => AddNumberFormat(formatIndexInFile, formatString);

    protected override string? ResolveSharedString(uint index) =>
        index < (uint)SharedStringCount ? Helpers.ConvertEscapeChars(GetSharedString((int)index)) : null;

    private void ReadWorkbook() => _zipWorker.LoadWorkbook(this);

    private void ReadSharedStrings()
    {
        _zipWorker.LoadSharedStrings(_stringStore);
        _stringStore.Seal();
    }

    private void ReadStyles() => _zipWorker.LoadStyles(this);
}
