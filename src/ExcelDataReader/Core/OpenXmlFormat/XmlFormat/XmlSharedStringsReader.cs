using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal sealed class XmlSharedStringsReader : XmlRecordReader
    {
        private const string ElementSst = "sst";
        private const string ElementStringItem = "si";

        public XmlSharedStringsReader(XmlReader reader)
            : base(reader)
        {
        }

        protected override IEnumerable<Record> ReadOverride(XmlProperNamespaces properNamespaces)
        {
            if (!Reader.IsStartElement(ElementSst, properNamespaces.NsSpreadsheetMl))
            {
                yield break;
            }

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(ElementStringItem, properNamespaces.NsSpreadsheetMl))
                {
                    var value = StringHelper.ReadStringItem(Reader, properNamespaces.NsSpreadsheetMl);
                    yield return new SharedStringRecord(value);
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }
        }
    }
}
