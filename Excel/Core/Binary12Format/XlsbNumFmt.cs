using System;
using System.Collections.Generic;
using System.Text;
using Excel.Core.OpenXmlFormat;

namespace Excel.Core.Binary12Format
{
	internal class XlsbNumFmt : XlsxNumFmt
	{
		public XlsbNumFmt(int id, string formatCode) : base(id, formatCode) { }
	}
}
