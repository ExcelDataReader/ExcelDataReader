namespace ExcelDataReader.Core.CompoundFormat
{
    internal enum STGTY : byte
    {
#pragma warning disable CA1712 // Do not prefix enum values with type name
        STGTY_INVALID = 0,
        STGTY_STORAGE = 1,
        STGTY_STREAM = 2,
        STGTY_LOCKBYTES = 3,
        STGTY_PROPERTY = 4,
        STGTY_ROOT = 5
#pragma warning restore CA1712 // Do not prefix enum values with type name
    }

    internal enum DECOLOR : byte
    {
        DE_RED = 0,
        DE_BLACK = 1
    }

    internal enum FATMARKERS : uint
    {
        FAT_EndOfChain = 0xFFFFFFFE,
        FAT_FreeSpace = 0xFFFFFFFF,
        FAT_FatSector = 0xFFFFFFFD,
        FAT_DifSector = 0xFFFFFFFC
    }
}
