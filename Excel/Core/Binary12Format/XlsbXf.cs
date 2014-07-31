using System;
using System.Collections.Generic;
using System.Text;
using Excel.Core.OpenXmlFormat;

namespace Excel.Core.Binary12Format
{
	internal class XlsbXf : XlsxXf
	{
		public XlsbXf(int id, int numFmtId, string applyNumberFormat) : base(id, numFmtId, applyNumberFormat) { }
	}
}