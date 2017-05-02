using System;

namespace ExcelDataReader.Core.OpenXmlFormat
{
	internal class XlsxDimension
	{
		public XlsxDimension(string value)
		{
			ParseDimensions(value);
		}

		public XlsxDimension(int rows, int cols)
		{
			this.FirstRow = 1;
			this.LastRow = rows;
			this.FirstCol = 1;
			this.LastCol = cols;

            IsRange = true;
        }

		private int _FirstRow;

		public int FirstRow
		{
			get { return _FirstRow; }
			set { _FirstRow = value; }
		}

		private int _LastRow;

		public int LastRow
		{
			get { return _LastRow; }
			set { _LastRow = value; }
		}

		private int _FirstCol;

		public int FirstCol
		{
			get { return _FirstCol; }
			set { _FirstCol = value; }
		}

		private int _LastCol;

		public int LastCol
		{
			get { return _LastCol; }
			set { _LastCol = value; }
		}

        private bool _IsRange;

        public bool IsRange
        {
            get { return _IsRange; }
            set { _IsRange = value; }
        }

        private void ParseDimensions(string value)
		{
			string[] parts = value.Split(':');

			int col;
			int row;

			ReferenceHelper.ParseReference(parts[0], out col, out row);
			FirstCol = col;
			FirstRow = row;

			if (parts.Length == 1)
			{
                LastCol = FirstCol;
				LastRow = FirstRow;
			}
			else
			{
			    ReferenceHelper.ParseReference(parts[1], out col, out row);
				LastCol = col;
				LastRow = row;

                IsRange = true;
            }

        }
    }
}
