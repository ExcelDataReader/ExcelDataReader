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
            string result = string.Empty;
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return result;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementT, NsSpreadsheetMl))
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    result += reader.ReadElementContentAsString();
                }
                else if (reader.IsStartElement(ElementR, NsSpreadsheetMl))
                {
                    result += ReadRichTextRun(reader);
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return result;
        }

        private static string ReadRichTextRun(XmlReader reader)
        {
            string result = string.Empty;
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return result;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementT, NsSpreadsheetMl))
                {
                    result += reader.ReadElementContentAsString();
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return result;
        }
    }
}
