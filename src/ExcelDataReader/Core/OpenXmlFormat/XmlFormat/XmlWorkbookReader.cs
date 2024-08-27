﻿using System.Globalization;
using System.Xml;

using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

internal sealed class XmlWorkbookReader(XmlReader reader, Dictionary<string, string> worksheetsRels) : XmlRecordReader(reader)
{
    private const string ElementWorkbook = "workbook";
    private const string ElementWorkbookProperties = "workbookPr";
    private const string ElementSheets = "sheets";
    private const string ElementSheet = "sheet";

    private const string AttributeSheetId = "sheetId";
    private const string AttributeVisibleState = "state";
    private const string AttributeName = "name";
    private const string AttributeRelationshipId = "id";

    private readonly Dictionary<string, string> _worksheetsRels = worksheetsRels;

    protected override IEnumerable<Record> ReadOverride()
    {
        if (!CheckStartElementAndApplyNamespaces(ElementWorkbook))
        {
            yield break;
        }

        if (!XmlReaderHelper.ReadFirstContent(Reader))
        {
            yield break;
        }

        while (!Reader.EOF)
        {
            if (Reader.IsStartElement(ElementWorkbookProperties, ProperNamespaces.NsSpreadsheetMl))
            {
                // Workbook VBA CodeName: reader.GetAttribute("codeName");
                bool date1904 = Reader.GetAttribute("date1904") is "1" or "true";
                yield return new WorkbookPrRecord(date1904);
                Reader.Skip();
            }
            else if (Reader.IsStartElement(ElementSheets, ProperNamespaces.NsSpreadsheetMl))
            {
                if (!XmlReaderHelper.ReadFirstContent(Reader))
                {
                    continue;
                }

                while (!Reader.EOF)
                {
                    if (Reader.IsStartElement(ElementSheet, ProperNamespaces.NsSpreadsheetMl))
                    {
                        var rid = Reader.GetAttribute(AttributeRelationshipId, ProperNamespaces.NsDocumentRelationship);
                        yield return new SheetRecord(
                            Reader.GetAttribute(AttributeName),
                            uint.Parse(Reader.GetAttribute(AttributeSheetId), CultureInfo.InvariantCulture),
                            rid,
                            Reader.GetAttribute(AttributeVisibleState),
                            rid != null && _worksheetsRels.TryGetValue(rid, out var path) ? path : null);
                        Reader.Skip();
                    }
                    else if (!XmlReaderHelper.SkipContent(Reader))
                    {
                        break;
                    }
                }
            }
            else if (Reader.IsStartElement("bookViews", ProperNamespaces.NsSpreadsheetMl))
            {
                if (!XmlReaderHelper.ReadFirstContent(Reader))
                {
                    continue;
                }

                while (!Reader.EOF)
                {
                    if (Reader.IsStartElement("workbookView", ProperNamespaces.NsSpreadsheetMl))
                    {
                        string activeTab = Reader.GetAttribute("activeTab");
                        int activeTabInt = int.TryParse(activeTab, out var result) ? result : 0;
                        yield return new WorkbookActRecord(activeTabInt);
                        Reader.Skip();
                    }
                    else if (!XmlReaderHelper.SkipContent(Reader))
                    {
                        break;
                    }
                }
            }
            else if (!XmlReaderHelper.SkipContent(Reader))
            {
                yield break;
            }
        }
    }

    private bool CheckStartElementAndApplyNamespaces(string element)
    {
        if (Reader.IsStartElement(element, ProperNamespaces.NsSpreadsheetMl))
        {
            return true;
        }

        return false;
    }
}
