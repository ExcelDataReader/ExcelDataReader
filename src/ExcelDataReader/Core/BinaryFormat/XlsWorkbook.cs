using System.Text;
using ExcelDataReader.Core.NumberFormat;
using ExcelDataReader.Core.OfficeCrypto;
using ExcelDataReader.Exceptions;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents Globals section of workbook.
/// </summary>
internal sealed class XlsWorkbook : CommonWorkbook, IWorkbook<XlsWorksheet>, IXlsWorkbookPartSink
{
    private readonly ISharedStringStore _stringStore = new SharedStringTable();

    internal XlsWorkbook(Stream stream, ExcelReaderConfiguration configuration)
    {
        Stream = stream;

        using var biffStream = new XlsBiffStream(stream, password: configuration.Password);
        if (biffStream.BiffVersion == 0)
            throw new ExcelReaderException(Errors.ErrorWorkbookGlobalsInvalidData);

        BiffVersion = biffStream.BiffVersion;
        SecretKey = biffStream.SecretKey;
        Encryption = biffStream.Encryption;
        Encoding = biffStream.BiffVersion == 8 ? Encoding.Unicode : configuration.FallbackEncoding;

        try
        {
            if (BiffVersion == 8 && configuration.SharedStringStorageMode != SharedStringStorageMode.Default)
                _stringStore = new AutoSharedStringStore(configuration);
            biffStream.SharedStringStore = _stringStore;

            switch (biffStream.BiffType)
            {
                case BIFFTYPE.WorkbookGlobals:
                    new XlsWorkbookPartReader(biffStream, this).Load();
                    _stringStore.Seal();
                    break;
                case BIFFTYPE.Worksheet:
                    // set up 'virtual' bound sheet pointing at this
                    Sheets.Add(new XlsBiffBoundSheet(0, XlsBiffBoundSheet.SheetType.Worksheet, XlsBiffBoundSheet.SheetVisibility.Visible, "Sheet"));
                    break;
                case BIFFTYPE.MacroSheet:
                    // set up 'virtual' bound sheet pointing at this
                    Sheets.Add(new XlsBiffBoundSheet(0, XlsBiffBoundSheet.SheetType.MacroSheet, XlsBiffBoundSheet.SheetVisibility.Visible, "Sheet"));
                    break;
                default:
                    throw new ExcelReaderException(Errors.ErrorWorkbookGlobalsInvalidData);
            }
        }
        catch (Exception exception)
        {
            ResourceCleanup.DisposeAll(exception, this);
            throw;
        }
    }

    public Stream Stream { get; }

    public int BiffVersion { get; }

    public byte[]? SecretKey { get; }

    public EncryptionInfo? Encryption { get; }

    public Encoding Encoding { get; private set; }

    Encoding IXlsWorkbookPartSink.Encoding => Encoding;

    public List<XlsBiffFont> Fonts { get; } = [];

    public List<XlsBiffBoundSheet> Sheets { get; } = [];

    /// <summary>
    /// Gets the shared string table of the workbook.
    /// </summary>
    public XlsBiffSST? SST { get; private set; }

    /// <summary>
    /// Gets a value indicating whether the workbook uses the 1904 date system.
    /// </summary>
    public bool IsDate1904 { get; private set; }

    public int ResultsCount => Sheets.Count;

    public int ActiveSheet { get; private set; }

    public static bool IsRawBiffStream(byte[] bytes)
    {
        if (bytes.Length < 8)
        {
            throw new ArgumentException("Needs at least 8 bytes to probe", nameof(bytes));
        }

        ushort rid = BitConverter.ToUInt16(bytes, 0);
        ushort size = BitConverter.ToUInt16(bytes, 2);
        ushort bofVersion = BitConverter.ToUInt16(bytes, 4);
        ushort type = BitConverter.ToUInt16(bytes, 6);

        switch (rid)
        {
            case 0x0009: // BIFF2
                if (size != 4)
                    return false;
                if (type != 0x10 && type != 0x20 && type != 0x40)
                    return false;
                return true;
            case 0x0209: // BIFF3
            case 0x0409: // BIFF4
                if (size != 6)
                    return false;
                if (type != 0x10 && type != 0x20 && type != 0x40 && type != 0x0100)
                    return false;
                /* removed this additional check to keep the probe at 8 bytes
                ushort notUsed = BitConverter.ToUInt16(bytes, 8);
                if (notUsed != 0x00)
                    return false;*/
                return true;
            case 0x0809: // BIFF5/BIFF8
                if (size < 4)
                    return false;
                if (bofVersion != 0 && bofVersion != 0x0200 && bofVersion != 0x0300 && bofVersion != 0x0400 && bofVersion != 0x0500 && bofVersion != 0x600)
                    return false;
                if (type != 0x5 && type != 0x6 && type != 0x10 && type != 0x20 && type != 0x40 && type != 0x0100)
                    return false;
                /* removed this additional check to keep the probe at 8 bytes
                ushort identifier = BitConverter.ToUInt16(bytes, 10);
                if (identifier == 0)
                    return false;*/
                return true;
        }

        return false;
    }

    public IEnumerable<XlsWorksheet> ReadWorksheets()
    {
        for (var i = 0; i < Sheets.Count; ++i)
        {
            yield return new XlsWorksheet(this, Sheets[i], Stream, SinglePassMode);
        }
    }

    public void Dispose()
    {
        ResourceCleanup.DisposeAll(null, _stringStore, Stream);
    }

    void IXlsWorkbookPartSink.AddSheet(XlsBiffBoundSheet sheet) => Sheets.Add(sheet);

    void IXlsWorkbookPartSink.AddFont(XlsBiffFont font) => Fonts.Add(font);

    void IXlsWorkbookPartSink.AddXf(XlsBiffXF xf) => AddXf(xf);

    void IXlsWorkbookPartSink.AddNumberFormat(int index, string format) =>
        Formats.Add(index, new NumberFormatString(format));

    void IXlsWorkbookPartSink.SetCodePage(ushort codePage) =>
        Encoding = EncodingHelper.GetEncoding(codePage) ?? Encoding;

    void IXlsWorkbookPartSink.SetDate1904(bool date1904) => IsDate1904 = date1904;

    void IXlsWorkbookPartSink.SetActiveSheet(int activeSheet) => ActiveSheet = activeSheet;

    void IXlsWorkbookPartSink.SetSharedStringTable(XlsBiffSST sharedStrings) => SST = sharedStrings;

    internal void AddXf(XlsBiffXF xf)
    {
        var extendedFormat = new ExtendedFormat(xf.ParentCellStyleXf, xf.Font, xf.Format, xf.IsLocked, xf.IsHidden, xf.IndentLevel, xf.HorizontalAlignment, xf.VerticalAlignment);

        // The workbook holds two kinds of XF records: Cell XFs, and Cell Style XFs.
        // In the binary XLS format, both kinds of XF records are saved in a single list,
        // whereas the XLSX format has two separate lists - like the CommonWorkbook internals.
        // The Cell XFs hold indexes into the Cell Style XF list, so adding the XF in both lists 
        // here to keep the indexes the same.
        ExtendedFormats.Add(extendedFormat);
        CellStyleExtendedFormats.Add(extendedFormat);
    }

    protected override string? ResolveSharedString(uint index)
    {
        if (SST == null)
            throw new ExcelReaderException(Errors.ErrorWorkbookGlobalsInvalidData);
        return index < (uint)_stringStore.Count ? _stringStore.GetString((int)index) : null;
    }
}
