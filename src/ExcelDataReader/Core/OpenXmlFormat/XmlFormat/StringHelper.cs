using System.Text;
using System.Xml;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

internal static class StringHelper
{       
    private const string ElementT = "t";
    private const string ElementR = "r";

    // https://www.w3.org/TR/REC-xml#NT-S
    private static readonly char[] WhitespaceChars = [' ', '\t', '\n', '\r'];

    public static string ReadStringItem(XmlReader reader, string nsSpreadsheetMl)
    {
        if (!XmlReaderHelper.ReadFirstContent(reader))
        {
            return string.Empty;
        }
        
        string? result = null;
        StringBuilder? sb = null;
        while (!reader.EOF)
        {
            if (reader.IsStartElement(ElementT, nsSpreadsheetMl))
            {
                // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                AppendElement(reader, ref result, ref sb);
            }
            else if (reader.IsStartElement(ElementR, nsSpreadsheetMl))
            {
                ReadRichTextRun(reader, ref result, ref sb, nsSpreadsheetMl);
            }
            else if (!XmlReaderHelper.SkipContent(reader))
            {
                break;
            }
        }

        return sb?.ToString() ?? result ?? string.Empty;
    }

    private static void AppendFragment(string fragment, ref string? result, ref StringBuilder? sb)
    {
        if (fragment.Length == 0)
            return;

        if (result == null)
            result = fragment;
        else
            (sb ??= new StringBuilder(result)).Append(fragment);
    }

    private static void AppendElement(XmlReader reader, ref string? result, ref StringBuilder? sb)
    {
        AppendFragment(ReadElementContent(reader), ref result, ref sb);
    }

    private static void ReadRichTextRun(XmlReader reader, ref string? result, ref StringBuilder? sb, string nsSpreadsheetMl)
    {
        if (!XmlReaderHelper.ReadFirstContent(reader))
        {
            return;
        }

        while (!reader.EOF)
        {
            if (reader.IsStartElement(ElementT, nsSpreadsheetMl))
            {
                AppendElement(reader, ref result, ref sb);
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
            return reader.ReadElementContentAsString().Trim(WhitespaceChars);
    }
}
