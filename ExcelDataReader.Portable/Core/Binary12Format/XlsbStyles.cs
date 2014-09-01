using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Core.Binary12Format
{
	internal class XlsbStyles
	{
		public XlsbStyles()
		{
			_cellXfs = new List<XlsbXf>();
			_NumFmts = new List<XlsbNumFmt>();
		}

		private List<XlsbXf> _cellXfs;

		public List<XlsbXf> CellXfs
		{
			get { return _cellXfs; }
			set { _cellXfs = value; }
		}

		private List<XlsbNumFmt> _NumFmts;

		public List<XlsbNumFmt> NumFmts
		{
			get { return _NumFmts; }
			set { _NumFmts = value; }
		}
	}
}
