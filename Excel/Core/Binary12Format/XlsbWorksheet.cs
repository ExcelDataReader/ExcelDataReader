using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Core.Binary12Format
{
	internal class XlsbWorksheet
	{
		private XlsbDimension _dimension;

		public XlsbDimension Dimension
		{
			get { return _dimension; }
			set { _dimension = value; }
		}

		public int ColumnsCount
		{
			get
			{
				return _dimension == null ? -1 : _dimension.LastCol - _dimension.FirstCol + 1;
			}
		}

		public int RowsCount
		{
			get
			{
				return _dimension == null ? -1 : _dimension.LastRow - _dimension.FirstRow + 1;
			}
		}

		private string _Name;

		public string Name
		{
			get { return _Name; }
		}

		private int _id;

		public int Id
		{
			get { return _id; }
		}

		public XlsbWorksheet(string name, int id)
		{
			_Name = name;
			_id = id;
		}
	}
}
