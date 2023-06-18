using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal static class StringHelper
    {       
        private const string ElementT = "t";
        private const string ElementR = "r";

        public static string ReadStringItem(XmlReader reader, string nsSpreadsheetMl)
        {
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return string.Empty;
            }

            StringBuilder sb = new();
            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementT, nsSpreadsheetMl))
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    sb.Append(ReadElementContent(reader));
                }
                else if (reader.IsStartElement(ElementR, nsSpreadsheetMl))
                {
                    ReadRichTextRun(reader, sb, nsSpreadsheetMl);
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return sb.ToString();
        }

        private static void ReadRichTextRun(XmlReader reader, StringBuilder sb, string nsSpreadsheetMl)
        {
            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(ElementT, nsSpreadsheetMl))
                {
                    sb.Append(ReadElementContent(reader));
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }
        }

        private static string ReadElementContent(XmlReader reader)
        {
            if (reader.GetAttribute("xml:space") == "preserve")
                return reader.ReadElementContentAsString();
            else
                return reader.ReadElementContentAsString().Trim();
        }
    }
}
