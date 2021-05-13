using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal static class StringHelper
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private const string ElementT = "t";
        private const string ElementR = "r";

        public static string ReadStringItem(XmlReader reader)
        {
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return string.Empty;
            }

            StringBuilder sb = new StringBuilder();
            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementT, NsSpreadsheetMl))
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    sb.Append(reader.ReadElementContentAsString());
                }
                else if (reader.IsStartElement(ElementR, NsSpreadsheetMl))
                {
                    ReadRichTextRun(reader, sb);
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return sb.ToString();
        }

        private static void ReadRichTextRun(XmlReader reader, StringBuilder sb)
        {
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementT, NsSpreadsheetMl))
                {
                    sb.Append(reader.ReadElementContentAsString());
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }
        }
    }
}
