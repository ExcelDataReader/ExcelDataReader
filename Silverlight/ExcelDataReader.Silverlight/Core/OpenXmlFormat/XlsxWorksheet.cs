namespace ExcelDataReader.Silverlight.Core.OpenXmlFormat
{
	internal class XlsxWorksheet
	{
		public const string N_dimension = "dimension";
		public const string N_row = "row";
		public const string N_c = "c";
		public const string N_v = "v";
		public const string A_ref = "ref";
		public const string A_r = "r";
		public const string A_t = "t";
		public const string A_s = "s";

		private XlsxDimension _dimension;

		public XlsxDimension Dimension
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

		private string _rid;

		public string RID
		{
			get
			{
				return _rid;
			}
			set
			{
				_rid = value;
			}
		}

		private string _path;

		public string Path
		{
			get
			{
				return _path;
			}
			set
			{
				_path = value;
			}
		}

		public XlsxWorksheet(string name, int id, string rid)
		{
			_Name = name;
			_id = id;
			_rid = rid;
		}

	}
}
