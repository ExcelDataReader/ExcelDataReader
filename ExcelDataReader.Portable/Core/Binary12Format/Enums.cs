namespace Excel.Core.Binary12Format
{
	/// <summary>
	/// BIFF 12 Record Types
	/// </summary>
	internal enum BIFF12 : ushort
	{
		/// <summary>
		/// 
		/// </summary>
		DEFINEDNAME = 0x27,
		/// <summary>
		/// 
		/// </summary>
		FILEVERSION = 0x0180,
		/// <summary>
		/// 
		/// </summary>
		WORKBOOK = 0x0183,
		/// <summary>
		/// 
		/// </summary>
		WORKBOOK_END = 0x0184,
		/// <summary>
		/// 
		/// </summary>
		BOOKVIEWS = 0x0187,
		/// <summary>
		/// 
		/// </summary>
		BOOKVIEWS_END = 0x0188,
		/// <summary>
		/// 
		/// </summary>
		SHEETS = 0x018F,
		/// <summary>
		/// 
		/// </summary>
		SHEETS_END = 0x0190,
		/// <summary>
		/// 
		/// </summary>
		WORKBOOKPR = 0x0199,
		/// <summary>
		/// 
		/// </summary>
		SHEET = 0x019C,
		/// <summary>
		/// 
		/// </summary>
		CALCPR = 0x019D,
		/// <summary>
		/// 
		/// </summary>
		WORKBOOKVIEW = 0x019E,
		/// <summary>
		/// 
		/// </summary>
		EXTERNALREFERENCES = 0x02E1,
		/// <summary>
		/// 
		/// </summary>
		EXTERNALREFERENCES_END = 0x02E2,
		/// <summary>
		/// 
		/// </summary>
		EXTERNALREFERENCE = 0x02E3,
		/// <summary>
		/// 
		/// </summary>
		WEBPUBLISHING = 0x04A9,

		// Worksheet records
		/// <summary>
		/// Row info.
		/// </summary>
		ROW = 0x00,
		/// <summary>
		/// Empty cell.
		/// </summary>
		BLANK = 0x01,
		/// <summary>
		/// Single-precision float.
		/// </summary>
		NUM = 0x02,
		/// <summary>
		/// Error identifier.
		/// </summary>
		BOOLERR = 0x03,
		/// <summary>
		/// Boolean value.
		/// </summary>
		BOOL = 0x04,
		/// <summary>
		/// Double-precision float.
		/// </summary>
		FLOAT = 0x05,
		/// <summary>
		/// String (shared string index).
		/// </summary>
		STRING = 0x07,
		/// <summary>
		/// Formula returning a string (inline string).
		/// </summary>
		FORMULA_STRING = 0x08,
		/// <summary>
		/// Formula returning a double-precision float.
		/// </summary>
		FORMULA_FLOAT = 0x09,
		/// <summary>
		/// Formula returning a boolean.
		/// </summary>
		FORMULA_BOOL = 0x0A,
		/// <summary>
		/// Formula returning an error identifier.
		/// </summary>
		FORMULA_BOOLERR = 0x0B,
		/// <summary>
		/// Column info.
		/// </summary>
		COL = 0x3C,
		/// <summary>
		/// 
		/// </summary>
		WORKSHEET = 0x0181,
		/// <summary>
		/// 
		/// </summary>
		WORKSHEET_END = 0x0182,
		/// <summary>
		/// 
		/// </summary>
		SHEETVIEWS = 0x0185,
		/// <summary>
		/// 
		/// </summary>
		SHEETVIEWS_END = 0x0186,
		/// <summary>
		/// 
		/// </summary>
		SHEETVIEW = 0x0189,
		/// <summary>
		/// 
		/// </summary>
		SHEETVIEW_END = 0x018A,
		/// <summary>
		/// 
		/// </summary>
		SHEETDATA = 0x0191,
		/// <summary>
		/// 
		/// </summary>
		SHEETDATA_END = 0x0192,
		/// <summary>
		/// 
		/// </summary>
		SHEETPR = 0x0193,
		/// <summary>
		/// 
		/// </summary>
		DIMENSION = 0x0194,
		/// <summary>
		/// 
		/// </summary>
		SELECTION = 0x0198,
		/// <summary>
		/// 
		/// </summary>
		COLS = 0x0386,
		/// <summary>
		/// 
		/// </summary>
		COLS_END = 0x0387,
		/// <summary>
		/// 
		/// </summary>
		CONDITIONALFORMATTING = 0x03CD,
		/// <summary>
		/// 
		/// </summary>
		CONDITIONALFORMATTING_END = 0x03CE,
		/// <summary>
		/// 
		/// </summary>
		CFRULE = 0x03CF,
		/// <summary>
		/// 
		/// </summary>
		CFRULE_END = 0x03D0,
		/// <summary>
		/// 
		/// </summary>
		ICONSET = 0x03D1,
		/// <summary>
		/// 
		/// </summary>
		ICONSET_END = 0x03D2,
		/// <summary>
		/// 
		/// </summary>
		DATABAR = 0x03D3,
		/// <summary>
		/// 
		/// </summary>
		DATABAR_END = 0x03D4,
		/// <summary>
		/// 
		/// </summary>
		COLORSCALE = 0x03D5,
		/// <summary>
		/// 
		/// </summary>
		COLORSCALE_END = 0x03D6,
		/// <summary>
		/// 
		/// </summary>
		CFVO = 0x03D7,
		/// <summary>
		/// 
		/// </summary>
		PAGEMARGINS = 0x03DC,
		/// <summary>
		/// 
		/// </summary>
		PRINTOPTIONS = 0x03DD,
		/// <summary>
		/// 
		/// </summary>
		PAGESETUP = 0x03DE,
		/// <summary>
		/// 
		/// </summary>
		HEADERFOOTER = 0x03DF,
		/// <summary>
		/// 
		/// </summary>
		SHEETFORMATPR = 0x03E5,
		/// <summary>
		/// 
		/// </summary>
		HYPERLINK = 0x03EE,
		/// <summary>
		/// 
		/// </summary>
		DRAWING = 0x04A6,
		/// <summary>
		/// 
		/// </summary>
		LEGACYDRAWING = 0x04A7,
		/// <summary>
		/// 
		/// </summary>
		COLOR = 0x04B4,
		/// <summary>
		/// 
		/// </summary>
		OLEOBJECTS = 0x04FE,
		/// <summary>
		/// 
		/// </summary>
		OLEOBJECT = 0x04FF,
		/// <summary>
		/// 
		/// </summary>
		OLEOBJECTS_END = 0x0580,
		/// <summary>
		/// 
		/// </summary>
		TABLEPARTS = 0x0594,
		/// <summary>
		/// 
		/// </summary>
		TABLEPART = 0x0595,
		/// <summary>
		/// 
		/// </summary>
		TABLEPARTS_END = 0x0596,

		//SharedStrings records
		/// <summary>
		/// 
		/// </summary>
		SI = 0x13,
		/// <summary>
		/// 
		/// </summary>
		SST = 0x019F,
		/// <summary>
		/// 
		/// </summary>
		SST_END = 0x01A0,

		//Styles records
		/// <summary>
		/// 
		/// </summary>
		FONT = 0x2B,
		/// <summary>
		/// 
		/// </summary>
		FILL = 0x2D,
		/// <summary>
		/// 
		/// </summary>
		BORDER = 0x2E,
		/// <summary>
		/// 
		/// </summary>
		XF = 0x2F,
		/// <summary>
		/// 
		/// </summary>
		CELLSTYLE = 0x30,
		/// <summary>
		/// 
		/// </summary>
		STYLESHEET = 0x0296,
		/// <summary>
		/// 
		/// </summary>
		STYLESHEET_END = 0x0297,
		/// <summary>
		/// 
		/// </summary>
		COLORS = 0x03D9,
		/// <summary>
		/// 
		/// </summary>
		COLORS_END = 0x03DA,
		/// <summary>
		/// 
		/// </summary>
		DXFS = 0x03F9,
		/// <summary>
		/// 
		/// </summary>
		DXFS_END = 0x03FA,
		/// <summary>
		/// 
		/// </summary>
		TABLESTYLES = 0x03FC,
		/// <summary>
		/// 
		/// </summary>
		TABLESTYLES_END = 0x03FD,
		/// <summary>
		/// 
		/// </summary>
		FILLS = 0x04DB,
		/// <summary>
		/// 
		/// </summary>
		FILLS_END = 0x04DC,
		/// <summary>
		/// 
		/// </summary>
		FONTS = 0x04E3,
		/// <summary>
		/// 
		/// </summary>
		FONTS_END = 0x04E4,
		/// <summary>
		/// 
		/// </summary>
		BORDERS = 0x04E5,
		/// <summary>
		/// 
		/// </summary>
		BORDERS_END = 0x04E6,
		/// <summary>
		/// 
		/// </summary>
		CELLXFS = 0x04E9,
		/// <summary>
		/// 
		/// </summary>
		CELLXFS_END = 0x04EA,
		/// <summary>
		/// 
		/// </summary>
		CELLSTYLES = 0x04EB,
		/// <summary>
		/// 
		/// </summary>
		CELLSTYLES_END = 0x04EC,
		/// <summary>
		/// 
		/// </summary>
		CELLSTYLEXFS = 0x04F2,
		/// <summary>
		/// 
		/// </summary>
		CELLSTYLEXFS_END = 0x04F3,

		//Comment records
		/// <summary>
		/// 
		/// </summary>
		COMMENTS = 0x04F4,
		/// <summary>
		/// 
		/// </summary>
		COMMENTS_END = 0x04F5,
		/// <summary>
		/// 
		/// </summary>
		AUTHORS = 0x04F6,
		/// <summary>
		/// 
		/// </summary>
		AUTHORS_END = 0x04F7,
		/// <summary>
		/// 
		/// </summary>
		AUTHOR = 0x04F8,
		/// <summary>
		/// 
		/// </summary>
		COMMENTLIST = 0x04F9,
		/// <summary>
		/// 
		/// </summary>
		COMMENTLIST_END = 0x04FA,
		/// <summary>
		/// 
		/// </summary>
		COMMENT = 0x04FB,
		/// <summary>
		/// 
		/// </summary>
		COMMENT_END = 0x04FC,
		/// <summary>
		/// 
		/// </summary>
		TEXT = 0x04FD,

		//Table records
		/// <summary>
		/// 
		/// </summary>
		AUTOFILTER = 0x01A1,
		/// <summary>
		/// 
		/// </summary>
		AUTOFILTER_END = 0x01A2,
		/// <summary>
		/// 
		/// </summary>
		FILTERCOLUMN = 0x01A3,
		/// <summary>
		/// 
		/// </summary>
		FILTERCOLUMN_END = 0x01A4,
		/// <summary>
		/// 
		/// </summary>
		FILTERS = 0x01A5,
		/// <summary>
		/// 
		/// </summary>
		FILTERS_END = 0x01A6,
		/// <summary>
		/// 
		/// </summary>
		FILTER = 0x01A7,
		/// <summary>
		/// 
		/// </summary>
		TABLE = 0x02D7,
		/// <summary>
		/// 
		/// </summary>
		TABLE_END = 0x02D8,
		/// <summary>
		/// 
		/// </summary>
		TABLECOLUMNS = 0x02D9,
		/// <summary>
		/// 
		/// </summary>
		TABLECOLUMNS_END = 0x02DA,
		/// <summary>
		/// 
		/// </summary>
		TABLECOLUMN = 0x02DB,
		/// <summary>
		/// 
		/// </summary>
		TABLECOLUMN_END = 0x02DC,
		/// <summary>
		/// 
		/// </summary>
		TABLESTYLEINFO = 0x0481,
		/// <summary>
		/// 
		/// </summary>
		SORTSTATE = 0x0492,
		/// <summary>
		/// 
		/// </summary>
		SORTCONDITION = 0x0494,
		/// <summary>
		/// 
		/// </summary>
		SORTSTATE_END = 0x0495,

		//QueryTable records
		/// <summary>
		/// 
		/// </summary>
		QUERYTABLE = 0x03BF,
		/// <summary>
		/// 
		/// </summary>
		QUERYTABLE_END = 0x03C0,
		/// <summary>
		/// 
		/// </summary>
		QUERYTABLEREFRESH = 0x03C1,
		/// <summary>
		/// 
		/// </summary>
		QUERYTABLEREFRESH_END = 0x03C2,
		/// <summary>
		/// 
		/// </summary>
		QUERYTABLEFIELDS = 0x03C7,
		/// <summary>
		/// 
		/// </summary>
		QUERYTABLEFIELDS_END = 0x03C8,
		/// <summary>
		/// 
		/// </summary>
		QUERYTABLEFIELD = 0x03C9,
		/// <summary>
		/// 
		/// </summary>
		QUERYTABLEFIELD_END = 0x03CA,

		//Connection records
		/// <summary>
		/// 
		/// </summary>
		CONNECTIONS = 0x03AD,
		/// <summary>
		/// 
		/// </summary>
		CONNECTIONS_END = 0x03AE,
		/// <summary>
		/// 
		/// </summary>
		CONNECTION = 0x01C9,
		/// <summary>
		/// 
		/// </summary>
		CONNECTION_END = 0x01CA,
		/// <summary>
		/// 
		/// </summary>
		DBPR = 0x01CB,
		/// <summary>
		/// 
		/// </summary>
		DBPR_END = 0x01CC,

		/// <summary>
		/// 
		/// </summary>
		UNKNOWN = 0x0
	}
}
