﻿using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

internal sealed class XmlSharedStringsReader(XmlReader reader) : XmlRecordReader(reader)
{
    private const string ElementSst = "sst";
    private const string ElementStringItem = "si";

    protected override IEnumerable<Record> ReadOverride()
    {
        if (!Reader.IsStartElement(ElementSst, ProperNamespaces.NsSpreadsheetMl))
        {
            yield break;
        }

        if (!XmlReaderHelper.ReadFirstContent(Reader))
        {
            yield break;
        }

        while (!Reader.EOF)
        {
            if (Reader.IsStartElement(ElementStringItem, ProperNamespaces.NsSpreadsheetMl))
            {
                var value = StringHelper.ReadStringItem(Reader, ProperNamespaces.NsSpreadsheetMl);
                yield return new SharedStringRecord(value);
            }
            else if (!XmlReaderHelper.SkipContent(Reader))
            {
                break;
            }
        }
    }
}
