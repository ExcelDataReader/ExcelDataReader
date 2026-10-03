using System.Globalization;
using System.IO.Compression;
using System.Text;

namespace ExcelDataReader.TestFixtures;

internal static class AllocationTestWorkbook
{
    public static byte[] CreateXlsx(string[] items, bool shared)
    {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            const string workbook = """
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets><sheet name="Sheet1" sheetId="1" r:id="sheet1"/></sheets>
                </workbook>
                """;
            WriteEntry(archive, "xl/workbook.xml", workbook);
            const string relationships = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="sheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="sst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
                </Relationships>
                """;
            WriteEntry(archive, "xl/_rels/workbook.xml.rels", relationships);
            var rows = new StringBuilder();
            var strings = new StringBuilder();
            for (int i = 0; i < items.Length; i++)
            {
                string row = (i + 1).ToString(CultureInfo.InvariantCulture);
                rows.Append("<row r=\"").Append(row).Append("\"><c r=\"A").Append(row);
                if (shared)
                {
                    rows.Append("\" t=\"s\"><v>").Append(i.ToString(CultureInfo.InvariantCulture)).Append("</v></c></row>");
                    strings.Append("<si>").Append(items[i]).Append("</si>");
                }
                else
                {
                    rows.Append("\" t=\"inlineStr\"><is>").Append(items[i]).Append("</is></c></row>");
                }
            }

            WriteEntry(archive, "xl/worksheets/sheet1.xml", "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" + rows + "</sheetData></worksheet>");
            if (shared)
            {
                WriteEntry(archive, "xl/sharedStrings.xml", "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" + strings + "</sst>");
            }
        }

        return buffer.ToArray();
    }

    public static byte[] CreateSpreadsheetXml(string[] data)
    {
        var xml = new StringBuilder("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"><Worksheet ss:Name=\"Sheet1\"><Table>");
        foreach (string item in data)
            xml.Append("<Row><Cell>").Append(item).Append("</Cell></Row>");
        xml.Append("</Table></Worksheet></Workbook>");
        return Encoding.UTF8.GetBytes(xml.ToString());
    }

    private static void WriteEntry(ZipArchive archive, string path, string text)
    {
        using var writer = new StreamWriter(archive.CreateEntry(path).Open(), new UTF8Encoding(false));
        writer.Write(text);
    }
}
