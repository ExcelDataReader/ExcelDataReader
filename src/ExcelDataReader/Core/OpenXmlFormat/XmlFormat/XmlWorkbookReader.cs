using System.Globalization;
using System.Xml;

using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

internal sealed class XmlWorkbookReader(XmlReader reader, Dictionary<string, string> worksheetsRels)
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
    private readonly string _nsSpreadsheetMl = XmlNamespaces.GetSpreadsheetMlNamespace(reader);

    public void Load(IWorkbookPartSink sink)
    {
        if (!reader.IsStartElement(ElementWorkbook, _nsSpreadsheetMl))
        {
            return;
        }

        if (!XmlReaderHelper.ReadFirstContent(reader))
        {
            return;
        }

        string nsDocumentRelationship = _nsSpreadsheetMl == XmlNamespaces.StrictNsSpreadsheetMl
            ? XmlNamespaces.StrictNsDocumentRelationship
            : XmlNamespaces.NsDocumentRelationship;
        while (!reader.EOF)
        {
            if (reader.IsStartElement(ElementWorkbookProperties, _nsSpreadsheetMl))
            {
                // Workbook VBA CodeName: reader.GetAttribute("codeName");
                bool date1904 = reader.GetAttribute("date1904") is "1" or "true";
                sink.SetDate1904(date1904);
                reader.Skip();
            }
            else if (reader.IsStartElement(ElementSheets, _nsSpreadsheetMl))
            {
                if (!XmlReaderHelper.ReadFirstContent(reader))
                {
                    continue;
                }

                while (!reader.EOF)
                {
                    if (reader.IsStartElement(ElementSheet, _nsSpreadsheetMl))
                    {
                        var rid = reader.GetAttribute(AttributeRelationshipId, nsDocumentRelationship);
                        var sheetId = reader.GetAttribute(AttributeSheetId);
                        sink.AddSheet(new SheetRecord(
                            reader.GetAttribute(AttributeName),
                            uint.TryParse(sheetId, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedSheetId) ? parsedSheetId : 0,
                            rid,
                            reader.GetAttribute(AttributeVisibleState),
                            rid != null ? _worksheetsRels.GetValueOrDefault(rid) : null));
                        reader.Skip();
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }
            else if (reader.IsStartElement("bookViews", _nsSpreadsheetMl))
            {
                if (!XmlReaderHelper.ReadFirstContent(reader))
                {
                    continue;
                }

                while (!reader.EOF)
                {
                    if (reader.IsStartElement("workbookView", _nsSpreadsheetMl))
                    {
                        string? activeTab = reader.GetAttribute("activeTab");
                        int activeTabInt = int.TryParse(activeTab, out var result) ? result : 0;
                        sink.SetActiveSheet(activeTabInt);
                        reader.Skip();
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }
            else if (!XmlReaderHelper.SkipContent(reader))
            {
                return;
            }
        }
    }
}
