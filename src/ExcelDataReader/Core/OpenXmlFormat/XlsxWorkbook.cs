using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat;

internal sealed class XlsxWorkbook : CommonWorkbook, IWorkbook<XlsxWorksheet>
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

    protected override string? ResolveSharedString(uint index) =>
        index < (uint)SharedStringCount ? Helpers.ConvertEscapeChars(GetSharedString((int)index)) : null;

    private void ReadWorkbook()
    {
        using var reader = _zipWorker.GetWorkbookReader();
        if (reader == null)
            return;

        while (reader.Read() is { } record)
        {                
            switch (record)
            {
                case WorkbookPrRecord pr:
                    IsDate1904 = pr.Date1904;
                    break;
                case SheetRecord sheet:
                    Sheets.Add(sheet);
                    break;
                case WorkbookActRecord activeSheet:
                    ActiveSheet = activeSheet.ActiveSheet;
                    break;
            }
        }
    }

    private void ReadSharedStrings()
    {
        _zipWorker.LoadSharedStrings(_stringStore);
        _stringStore.Seal();
    }

    private void ReadStyles()
    {
        using var reader = _zipWorker.GetStylesReader();
        if (reader == null)
            return;

        while (reader.Read() is { } record)
        {
            switch (record)
            {
                case ExtendedFormatRecord xf:
                    ExtendedFormats.Add(xf.ExtendedFormat);
                    break;
                case CellStyleExtendedFormatRecord csxf:
                    CellStyleExtendedFormats.Add(csxf.ExtendedFormat);
                    break;
                case NumberFormatRecord nf:
                    AddNumberFormat(nf.FormatIndexInFile, nf.FormatString);
                    break;
            }
        }
    }
}
