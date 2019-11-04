using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal sealed class XmlWorkbookReader : XmlRecordReader
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NsDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private const string ElementWorkbook = "workbook";
        private const string ElementWorkbookProperties = "workbookPr";
        private const string ElementSheets = "sheets";
        private const string ElementSheet = "sheet";

        private const string AttributeSheetId = "sheetId";
        private const string AttributeVisibleState = "state";
        private const string AttributeName = "name";
        private const string AttributeRelationshipId = "id";

        public XmlWorkbookReader(XmlReader reader)
            : base(reader)
        {
        }

        protected override IEnumerable<Record> ReadOverride()
        {
            if (!Reader.IsStartElement(ElementWorkbook, NsSpreadsheetMl))
            {
                yield break;
            }

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(ElementWorkbookProperties, NsSpreadsheetMl))
                {
                    // Workbook VBA CodeName: reader.GetAttribute("codeName");
                    bool date1904 = Reader.GetAttribute("date1904") == "1";
                    yield return new WorkbookPrRecord(date1904);
                    Reader.Skip();
                }
                else if (Reader.IsStartElement(ElementSheets, NsSpreadsheetMl))
                {
                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        continue;
                    }

                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(ElementSheet, NsSpreadsheetMl))
                        {
                            yield return new SheetRecord(
                                Reader.GetAttribute(AttributeName),
                                uint.Parse(Reader.GetAttribute(AttributeSheetId)),
                                Reader.GetAttribute(AttributeRelationshipId, NsDocumentRelationship),
                                Reader.GetAttribute(AttributeVisibleState));
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
    }
}
