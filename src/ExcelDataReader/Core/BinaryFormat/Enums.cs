// ReSharper disable InconsistentNaming
namespace ExcelDataReader.Core.BinaryFormat
{
    internal enum BIFFTYPE : ushort
    {
        WorkbookGlobals = 0x0005,
        VBModule = 0x0006,
        Worksheet = 0x0010,
        Chart = 0x0020,
#pragma warning disable SA1300 // Element must begin with upper-case letter
        v4MacroSheet = 0x0040,
        v4WorkbookGlobals = 0x0100
#pragma warning restore SA1300 // Element must begin with upper-case letter
    }

    internal enum BIFFRECORDTYPE : ushort
    {
        INTERFACEHDR = 0x00E1,
        MMS = 0x00C1,
        MERGECELLS = 0x00E5, // Record containing list of merged cell ranges
        INTERFACEEND = 0x00E2,
        WRITEACCESS = 0x005C,
        CODEPAGE = 0x0042,
        DSF = 0x0161,
        TABID = 0x013D,
        FNGROUPCOUNT = 0x009C,
        FILEPASS = 0x002F,
        WINDOWPROTECT = 0x0019,
        PROTECT = 0x0012,
        PASSWORD = 0x0013,
        PROT4REV = 0x01AF,
        PROT4REVPASSWORD = 0x01BC,
        WINDOW1 = 0x003D,
        BACKUP = 0x0040,
        HIDEOBJ = 0x008D,
        RECORD1904 = 0x0022,
        REFRESHALL = 0x01B7,
        BOOKBOOL = 0x00DA,

        FONT = 0x0031, // Font record, BIFF2, 5 and later

        FONT_V34 = 0x0231, // Font record, BIFF3, 4

        FORMAT = 0x041E, // Format record, BIFF4 and later

        FORMAT_V23 = 0x001E, // Format record, BIFF2, 3

        XF = 0x00E0, // Extended format record, BIFF5 and later

        XF_V4 = 0x0443, // Extended format record, BIFF4

        XF_V3 = 0x0243, // Extended format record, BIFF3

        XF_V2 = 0x0043, // Extended format record, BIFF2

        IXFE = 0x0044, // Index to XF, BIFF2

        STYLE = 0x0293,
        BOUNDSHEET = 0x0085,
        COUNTRY = 0x008C,
        SST = 0x00FC, // Global string storage (for BIFF8)

        CONTINUE = 0x003C,
        EXTSST = 0x00FF,
        BOF = 0x0809, // BOF Id for BIFF5 and later

        BOF_V2 = 0x0009, // BOF Id for BIFF2

        BOF_V3 = 0x0209, // BOF Id for BIFF3

        BOF_V4 = 0x0409, // BOF Id for BIFF4

        EOF = 0x000A, // End of block started with BOF

        CALCCOUNT = 0x000C,
        CALCMODE = 0x000D,
        PRECISION = 0x000E,
        REFMODE = 0x000F,
        DELTA = 0x0010,
        ITERATION = 0x0011,
        SAVERECALC = 0x005F,
        PRINTHEADERS = 0x002A,
        PRINTGRIDLINES = 0x002B,
        GUTS = 0x0080,
        WSBOOL = 0x0081,
        GRIDSET = 0x0082,
        DEFAULTROWHEIGHT_V2 = 0x0025,
        DEFAULTROWHEIGHT = 0x0225,
        HEADER = 0x0014,
        FOOTER = 0x0015,
        HCENTER = 0x0083,
        VCENTER = 0x0084,
        PRINTSETUP = 0x00A1,
        DFAULTCOLWIDTH = 0x0055,
        DIMENSIONS = 0x0200, // Size of area used for data
        DIMENSIONS_V2 = 0x0000, // BIFF2

        ROW_V2 = 0x0008, // Row record
        ROW = 0x0208, // Row record

        WINDOW2 = 0x023E,
        SELECTION = 0x001D,
        INDEX = 0x020B, // Index record, unsure about signature

        DBCELL = 0x00D7, // DBCell record, unsure about signature

        BLANK = 0x0201, // Empty cell

        BLANK_OLD = 0x0001, // Empty cell, old format

        MULBLANK = 0x00BE, // Equivalent of up to 256 blank cells

        INTEGER = 0x0202, // Integer cell (0..65535)

        INTEGER_OLD = 0x0002, // Integer cell (0..65535), old format

        NUMBER = 0x0203, // Numeric cell

        NUMBER_OLD = 0x0003, // Numeric cell, old format

        LABEL = 0x0204, // String cell (up to 255 symbols)

        LABEL_OLD = 0x0004, // String cell (up to 255 symbols), old format

        LABELSST = 0x00FD, // String cell with value from SST (for BIFF8)

        FORMULA = 0x0006, // Formula cell, BIFF2, BIFF5-8

        FORMULA_V3 = 0x0206, // Formula cell, BIFF3

        FORMULA_V4 = 0x0406, // Formula cell, BIFF4

        BOOLERR = 0x0205, // Boolean or error cell

        BOOLERR_OLD = 0x0005, // Boolean or error cell, old format

        ARRAY = 0x0221, // Range of cells for multi-cell formula

        RK = 0x027E, // RK-format numeric cell

        MULRK = 0x00BD, // Equivalent of up to 256 RK cells

        RSTRING = 0x00D6, // Rich-formatted string cell

        SHAREDFMLA = 0x04BC, // One more formula optimization element

        SHAREDFMLA_OLD = 0x00BC, // One more formula optimization element, old format

        STRING = 0x0207, // And one more, for string formula results

        STRING_OLD = 0x0007, // Old string formula results

        CF = 0x01B1,
        CODENAME = 0x01BA,
        CONDFMT = 0x01B0,
        DCONBIN = 0x01B5,
        DV = 0x01BE,
        DVAL = 0x01B2,
        HLINK = 0x01B8,
        MSODRAWINGGROUP = 0x00EB,
        MSODRAWING = 0x00EC,
        MSODRAWINGSELECTION = 0x00ED,
        PARAMQRY = 0x00DC,
        QSI = 0x01AD,
        SUPBOOK = 0x01AE,
        SXDB = 0x00C6,
        SXDBEX = 0x0122,
        SXFDBTYPE = 0x01BB,
        SXRULE = 0x00F0,
        SXEX = 0x00F1,
        SXFILT = 0x00F2,
        SXNAME = 0x00F6,
        SXSELECT = 0x00F7,
        SXPAIR = 0x00F8,
        SXFMLA = 0x00F9,
        SXFORMAT = 0x00FB,
        SXFORMULA = 0x0103,
        SXVDEX = 0x0100,
        TXO = 0x01B6,
        USERBVIEW = 0x01A9,
        USERSVIEWBEGIN = 0x01AA,
        USERSVIEWEND = 0x01AB,
        USESELFS = 0x0160,
        XL5MODIFY = 0x0162,
        OBJ = 0x005D,
        NOTE = 0x001C,
        SXEXT = 0x00DC,
        VERTICALPAGEBREAKS = 0x001A,
        XCT = 0x0059,

        /// <summary>
        /// If present the Calculate Message was in the status bar when Excel saved the file.
        /// This occurs if the sheet changed, the Manual calculation option was on, and the Recalculate Before Save option was off.
        /// </summary>
        UNCALCED = 0x005E,
        QUICKTIP = 0x0800,
        COLINFO = 0x007D
    }

    internal enum FORMULAERROR : byte
    {
        NULL = 0x00, // #NULL!

        DIV0 = 0x07, // #DIV/0!

        VALUE = 0x0F, // #VALUE!

        REF = 0x17, // #REF!

        NAME = 0x1D, // #NAME?

        NUM = 0x24, // #NUM!

        NA = 0x2A, // #N/A
    }
}
