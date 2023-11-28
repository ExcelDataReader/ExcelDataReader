using System.Collections.Generic;

using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal sealed class XlsxWorkbook : CommonWorkbook, IWorkbook<XlsxWorksheet>
    {
        private readonly ZipWorker _zipWorker;
               
        public XlsxWorkbook(ZipWorker zipWorker)
        {
            _zipWorker = zipWorker;
            ReadWorkbook();
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
            using RecordReader reader = _zipWorker.GetWorkbookReader();            
            
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
                    default:
                        {
                            break;
                        }
                }
            }
        }
    }
}
