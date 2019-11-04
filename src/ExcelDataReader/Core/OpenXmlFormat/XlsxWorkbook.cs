using System;
using System.Collections.Generic;
using System.Xml;

using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxWorkbook : CommonWorkbook, IWorkbook<XlsxWorksheet>
    {
        private const string NsRelationship = "http://schemas.openxmlformats.org/package/2006/relationships";

        private const string ElementRelationship = "Relationship";
        private const string ElementRelationships = "Relationships";
        private const string AttributeId = "Id";
        private const string AttributeTarget = "Target";

        private readonly ZipWorker _zipWorker;

        public XlsxWorkbook(ZipWorker zipWorker)
        {
            _zipWorker = zipWorker;

            ReadWorkbook();
            ReadWorkbookRels();
            ReadSharedStrings();
            ReadStyles();
        }

        public List<SheetRecord> Sheets { get; } = new List<SheetRecord>();

        public XlsxSST SST { get; } = new XlsxSST();

        public bool IsDate1904 { get; private set; }

        public int ResultsCount => Sheets?.Count ?? -1;

        public IEnumerable<XlsxWorksheet> ReadWorksheets()
        {
            foreach (var sheet in Sheets)
            {
                yield return new XlsxWorksheet(_zipWorker, this, sheet);
            }
        }

        private void ReadWorkbook()
        {
            using var reader = _zipWorker.GetWorkbookReader();

            Record record;
            while ((record = reader.Read()) != null)
            {
                switch (record)
                {
                    case WorkbookPrRecord pr:
                        IsDate1904 = pr.Date1904;
                        break;
                    case SheetRecord sheet:
                        Sheets.Add(sheet);
                        break;
                }
            }
        }

        private void ReadWorkbookRels()
        {
            using var stream = _zipWorker.GetWorkbookRelsStream();
            if (stream == null)
            {
                return;
            }

            using XmlReader reader = XmlReader.Create(stream);
            ReadWorkbookRels(reader);
        }

        private void ReadWorkbookRels(XmlReader reader)
        {
            if (!reader.IsStartElement(ElementRelationships, NsRelationship))
            {
                return;
            }

            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementRelationship, NsRelationship))
                {
                    string rid = reader.GetAttribute(AttributeId);
                    foreach (var sheet in Sheets)
                    {
                        if (sheet.Rid == rid)
                        {
                            sheet.Path = reader.GetAttribute(AttributeTarget);
                            break;
                        }
                    }

                    reader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }
        }

        private void ReadSharedStrings()
        {
            using var reader = _zipWorker.GetSharedStringsReader();
            if (reader == null)
                return;

            Record record;
            while ((record = reader.Read()) != null)
            {
                switch (record)
                {
                    case SharedStringRecord pr:
                        SST.Add(pr.Value);
                        break;
                }
            }
        }

        private void ReadStyles()
        {
            using var reader = _zipWorker.GetStylesReader();
            if (reader == null)
                return;

            Record record;
            while ((record = reader.Read()) != null)
            {
                switch (record)
                {
                    case ExtendedFormatRecord xf:
                        ExtendedFormats.Add(xf.ExtendedFormat);
                        break;
                    case CellStyleExtendedFormatRecord csxf:
                        CellStyleExtendedFormats.Add(csxf.ExtendedFormat);
                        break;
                    case NumberFormatRecord nf:
                        AddNumberFormat(nf.FormatIndexInFile, nf.FormatString);
                        break;
                }
            }
        }
    }
}
