using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal class XlsxWorkbook : IWorkbook<XlsxWorksheet>
    {
        private const string ElementSheet = "sheet";
        private const string ElementT = "t";
        private const string ElementStringItem = "si";
        private const string ElementCellCrossReference = "cellXfs";
        private const string ElementNumberFormats = "numFmts";
        private const string ElementWorkbookProperties = "workbookPr";

        private const string AttributeSheetId = "sheetId";
        private const string AttributeVisibleState = "state";
        private const string AttributeName = "name";
        private const string AttributeRelationshipId = "r:id";

        private const string ElementRelationship = "Relationship";
        private const string AttributeId = "Id";
        private const string AttributeTarget = "Target";

        private readonly List<int> _defaultDateTimeStyles;
        private ZipWorker _zipWorker;

        public XlsxWorkbook(ZipWorker zipWorker)
        {
            _defaultDateTimeStyles = new List<int>(new[]
            {
                14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47
            });

            _zipWorker = zipWorker;

            using (var stream = _zipWorker.GetWorkbookStream())
            {
                Sheets = ReadWorkbook(stream);
            }

            using (var stream = _zipWorker.GetWorkbookRelsStream())
            {
                ReadWorkbookRels(stream, Sheets);
            }

            using (var stream = _zipWorker.GetSharedStringsStream())
            {
                SST = ReadSharedStrings(stream);
            }

            using (var stream = _zipWorker.GetStylesStream())
            {
                Styles = ReadStyles(stream);
            }

            CheckDateTimeNumFmts(Styles.NumFmts);
        }

        public List<XlsxBoundSheet> Sheets { get; }

        public XlsxSST SST { get; }

        public XlsxStyles Styles { get; }

        public Encoding Encoding => null;

        public bool IsDate1904 { get; private set; }

        public int ResultsCount => Sheets?.Count ?? -1;

        public IEnumerable<XlsxWorksheet> ReadWorksheets()
        {
            foreach (var sheet in Sheets)
            {
                yield return new XlsxWorksheet(_zipWorker, this, sheet);
            }
        }

        public bool IsDateTimeStyle(int styleId)
        {
            return _defaultDateTimeStyles.Contains(styleId);
        }

        private void CheckDateTimeNumFmts(List<XlsxNumFmt> list)
        {
            if (list.Count == 0)
                return;

            foreach (XlsxNumFmt numFmt in list)
            {
                if (string.IsNullOrEmpty(numFmt.FormatCode))
                    continue;
                string fc = numFmt.FormatCode.ToLower();

                int pos;
                while ((pos = fc.IndexOf('"')) > 0)
                {
                    int endPos = fc.IndexOf('"', pos + 1);

                    if (endPos > 0)
                        fc = fc.Remove(pos, endPos - pos + 1);
                }

                // it should only detect it as a date if it contains
                // dd mm mmm yy yyyy
                // h hh ss
                // AM PM
                // and only if these appear as "words" so either contained in [ ]
                // or delimted in someway
                // updated to not detect as date if format contains a #
                var formatReader = new FormatReader { FormatString = fc };
                if (formatReader.IsDateFormatString())
                {
                    _defaultDateTimeStyles.Add(numFmt.Id);
                }
            }
        }

        private List<XlsxBoundSheet> ReadWorkbook(Stream xmlFileStream)
        {
            var sheets = new List<XlsxBoundSheet>();

            using (XmlReader reader = XmlReader.Create(xmlFileStream))
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementWorkbookProperties)
                    {
                        IsDate1904 = reader.GetAttribute("date1904") == "1";
                    }
                    else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementSheet)
                    {
                        sheets.Add(new XlsxBoundSheet(
                            reader.GetAttribute(AttributeName),
                            int.Parse(reader.GetAttribute(AttributeSheetId)),
                            reader.GetAttribute(AttributeRelationshipId),
                            reader.GetAttribute(AttributeVisibleState)));
                    }
                }
            }

            return sheets;
        }

        private void ReadWorkbookRels(Stream xmlFileStream, List<XlsxBoundSheet> sheets)
        {
            using (XmlReader reader = XmlReader.Create(xmlFileStream))
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementRelationship)
                    {
                        string rid = reader.GetAttribute(AttributeId);

                        for (int i = 0; i < sheets.Count; i++)
                        {
                            var tempSheet = sheets[i];

                            if (tempSheet.Rid == rid)
                            {
                                tempSheet.Path = reader.GetAttribute(AttributeTarget);
                                sheets[i] = tempSheet;
                                break;
                            }
                        }
                    }
                }
            }
        }

        private XlsxSST ReadSharedStrings(Stream xmlFileStream)
        {
            if (xmlFileStream == null)
                return null;

            var sst = new XlsxSST();

            using (XmlReader reader = XmlReader.Create(xmlFileStream))
            {
                // Skip phonetic string data.
                bool bSkipPhonetic = false;

                // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                bool bAddStringItem = false;
                string sStringItem = string.Empty;

                while (reader.Read())
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementStringItem)
                    {
                        // Do not add the string item until the next string item is read.
                        if (bAddStringItem)
                        {
                            // Add the string item to XlsxSST.
                            sst.Add(sStringItem);
                        }
                        else
                        {
                            // Add the string items from here on.
                            bAddStringItem = true;
                        }

                        // Reset the string item.
                        sStringItem = string.Empty;
                    }
                    else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementT)
                    {
                        // Skip phonetic string data.
                        if (!bSkipPhonetic)
                        {
                            // Append to the string item.
                            sStringItem += reader.ReadElementContentAsString();
                        }
                    }

                    if (reader.LocalName == "rPh")
                    {
                        // Phonetic items represents pronunciation hints for some East Asian languages.
                        // In the file 'xl/sharedStrings.xml', the phonetic properties appear like:
                        // <si>
                        //  <t>(a japanese text in KANJI)</t>
                        //  <rPh sb="0" eb="1">
                        //      <t>(its pronounciation in KATAKANA)</t>
                        //  </rPh>
                        // </si>
                        if (reader.NodeType == XmlNodeType.Element)
                            bSkipPhonetic = true;
                        else if (reader.NodeType == XmlNodeType.EndElement)
                            bSkipPhonetic = false;
                    }
                }

                // Do not add the last string item unless we have read previous string items.
                if (bAddStringItem)
                {
                    // Add the string item to XlsxSST.
                    sst.Add(sStringItem);
                }
            }

            return sst;
        }

        private XlsxStyles ReadStyles(Stream xmlFileStream)
        {
            var styles = new XlsxStyles();

            if (xmlFileStream == null)
                return styles;

            bool rXlsxNumFmt = false;

            using (XmlReader reader = XmlReader.Create(xmlFileStream))
            {
                while (reader.Read())
                {
                    if (!rXlsxNumFmt && reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementNumberFormats)
                    {
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1)
                                break;

                            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxNumFmt.NNumFmt)
                            {
                                styles.NumFmts.Add(
                                    new XlsxNumFmt(
                                        int.Parse(reader.GetAttribute(XlsxNumFmt.ANumFmtId)),
                                        reader.GetAttribute(XlsxNumFmt.AFormatCode)));
                            }
                        }

                        rXlsxNumFmt = true;
                    }

                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == ElementCellCrossReference)
                    {
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1)
                                break;

                            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxXf.NXF)
                            {
                                var xfId = reader.GetAttribute(XlsxXf.AXFId);
                                var numFmtId = reader.GetAttribute(XlsxXf.ANumFmtId);

                                styles.CellXfs.Add(
                                    new XlsxXf(
                                        xfId == null ? -1 : int.Parse(xfId),
                                        numFmtId == null ? -1 : int.Parse(numFmtId),
                                        reader.GetAttribute(XlsxXf.AApplyNumberFormat)));
                            }
                        }

                        break;
                    }
                }
            }

            return styles;
        }
    }
}
