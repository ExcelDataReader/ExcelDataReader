using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents Sheet record in Workbook Globals.
/// </summary>
internal sealed class XlsSheetInfo
{
    private readonly XlsStringData _sheetName;

    internal XlsSheetInfo(XlsBiffFrame frame, int biffVersion)
    {
        StartOffset = frame.ReadUInt32(0x0);
        Type = (SheetType)frame.ReadByte(0x5);
        VisibleState = (SheetVisibility)frame.ReadByte(0x4);

        if (biffVersion == 8)
        {
            _sheetName = XlsStringData.Read(frame, 10, true, true);
        }
        else if (biffVersion == 5)
        {
            _sheetName = XlsStringData.Read(frame, 10, true, false);
        }
        else
        {
            throw new ArgumentException("Unexpected BIFF version " + biffVersion, nameof(biffVersion));
        }
    }

    internal XlsSheetInfo(uint startOffset, SheetType type, SheetVisibility visibleState, string name)
    {
        StartOffset = startOffset;
        Type = type;
        VisibleState = visibleState;
        _sheetName = new XlsStringData(name);
    }

    public enum SheetType : byte
    {
        Worksheet = 0x0,
        MacroSheet = 0x1,
        Chart = 0x2,

        // ReSharper disable once InconsistentNaming
        VBModule = 0x6
    }

    public enum SheetVisibility : byte
    {
        Visible = 0x0,
        Hidden = 0x1,
        VeryHidden = 0x2
    }

    /// <summary>
    /// Gets the worksheet data start offset.
    /// </summary>
    public uint StartOffset { get; }

    /// <summary>
    /// Gets the worksheet type.
    /// </summary>
    public SheetType Type { get; }

    /// <summary>
    /// Gets the visibility of the worksheet.
    /// </summary>
    public SheetVisibility VisibleState { get; }

    /// <summary>
    /// Gets the name of the worksheet.
    /// </summary>
    public string GetSheetName(Encoding encoding)
    {
        return _sheetName.GetValue(encoding);
    }
}