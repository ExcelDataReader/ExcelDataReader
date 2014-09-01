using System.Collections.Generic;

namespace ExcelDataReader.Portable.Core.OpenXmlFormat
{
	internal class XlsxStyles
	{
		public XlsxStyles()
		{
			_cellXfs = new List<XlsxXf>();
			_NumFmts = new List<XlsxNumFmt>();
		}

		private List<XlsxXf> _cellXfs;

		public List<XlsxXf> CellXfs
		{
			get { return _cellXfs; }
			set { _cellXfs = value; }
		}

		private List<XlsxNumFmt> _NumFmts;

		public List<XlsxNumFmt> NumFmts
		{
			get { return _NumFmts; }
			set { _NumFmts = value; }
		}
	}
}
