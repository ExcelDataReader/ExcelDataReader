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
        private const string ElementWorkbook = "workbook";
        private const string ElementWorkbookProperties = "workbookPr";
        private const string ElementSheets = "sheets";
        private const string ElementSheet = "sheet";

        private const string AttributeSheetId = "sheetId";
        private const string AttributeVisibleState = "state";
        private const string AttributeName = "name";
        private const string AttributeRelationshipId = "id";

        public XmlWorkbookReader(XmlReader reader, XmlProperNamespaces properNamespaces)
            : base(reader, properNamespaces)
        {
        }

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
                    bool date1904 = Reader.GetAttribute("date1904") == "1";
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
                            yield return new SheetRecord(
                                Reader.GetAttribute(AttributeName),
                                uint.Parse(Reader.GetAttribute(AttributeSheetId)),
                                Reader.GetAttribute(AttributeRelationshipId, ProperNamespaces.NsDocumentRelationship),
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

        private bool CheckStartElementAndApplyNamespaces(string element)
        {
            if (Reader.IsStartElement(element, XmlNamespaces.NsSpreadsheetMl))
            {
                return true;
            }

            if (Reader.IsStartElement(element, XmlNamespaces.StrictNsSpreadsheetMl))
            {
                ProperNamespaces.SetStrictNamespaces();
                return true;
            }

            return false;
        }
    }
}
