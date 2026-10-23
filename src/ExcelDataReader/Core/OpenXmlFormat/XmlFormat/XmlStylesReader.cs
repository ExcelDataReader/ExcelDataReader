using System.Globalization;
using System.Xml;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

internal sealed class XmlStylesReader(XmlReader reader)
{
    private const string ElementStyleSheet = "styleSheet";

    private const string ANumFmtId = "numFmtId";

    private const string ElementCellCrossReference = "cellXfs";
    private const string ElementCellStyleCrossReference = "cellStyleXfs";
    private const string NXF = "xf";
    private const string AXFId = "xfId";

    // private const string AApplyNumberFormat = "applyNumberFormat";
    // private const string AApplyAlignment = "applyAlignment";
    // private const string AApplyProtection = "applyProtection";
    private const string ElementNumberFormats = "numFmts";
    private const string NNumFmt = "numFmt";
    private const string AFormatCode = "formatCode";

    private const string NAlignment = "alignment";
    private const string AIndent = "indent";
    private const string AHorizontal = "horizontal";

    private const string AVertical = "vertical";

    private const string NProtection = "protection";
    private const string AHidden = "hidden";
    private const string ALocked = "locked";

    private readonly XmlProperNamespaces _properNamespaces = new(reader.IsStartElement() && reader.NamespaceURI == XmlNamespaces.StrictNsSpreadsheetMl);

    public void Load(IStylesPartSink sink)
    {
        if (!reader.IsStartElement(ElementStyleSheet, _properNamespaces.NsSpreadsheetMl))
        {
            return;
        }

        if (!XmlReaderHelper.ReadFirstContent(reader))
        {
            return;
        }

        while (!reader.EOF)
        {
            if (reader.IsStartElement(ElementCellCrossReference, _properNamespaces.NsSpreadsheetMl))
            {
                ReadCellXfs(_properNamespaces.NsSpreadsheetMl, sink.AddExtendedFormat);
            }
            else if (reader.IsStartElement(ElementCellStyleCrossReference, _properNamespaces.NsSpreadsheetMl))
            {
                ReadCellXfs(_properNamespaces.NsSpreadsheetMl, sink.AddCellStyleExtendedFormat);
            }
            else if (reader.IsStartElement(ElementNumberFormats, _properNamespaces.NsSpreadsheetMl))
            {
                if (!XmlReaderHelper.ReadFirstContent(reader))
                {
                    continue;
                }

                while (!reader.EOF)
                {
                    if (reader.IsStartElement(NNumFmt, _properNamespaces.NsSpreadsheetMl))
                    {
                        int.TryParse(reader.GetAttribute(ANumFmtId), NumberStyles.Integer, CultureInfo.InvariantCulture, out var numFmtId);
                        var formatCode = reader.GetAttribute(AFormatCode);

                        sink.AddNumberFormat(numFmtId, formatCode ?? string.Empty);
                        reader.Skip();
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }
            else if (!XmlReaderHelper.SkipContent(reader))
            {
                break;
            }
        }
    }

    private void ReadCellXfs(string nsSpreadsheetMl, Action<ExtendedFormat> add)
    {
        if (!XmlReaderHelper.ReadFirstContent(reader))
        {
            return;
        }

        while (!reader.EOF)
        {
            if (reader.IsStartElement(NXF, nsSpreadsheetMl))
            {
                int.TryParse(reader.GetAttribute(AXFId), NumberStyles.Integer, CultureInfo.InvariantCulture, out var xfId);
                int.TryParse(reader.GetAttribute(ANumFmtId), NumberStyles.Integer, CultureInfo.InvariantCulture, out var numFmtId);
                
                // var applyNumberFormat = reader.GetAttribute(AApplyNumberFormat) == "1";
                // var applyAlignment = reader.GetAttribute(AApplyAlignment) == "1";
                // var applyProtection = reader.GetAttribute(AApplyProtection) == "1";
                ReadAlignment(reader, nsSpreadsheetMl, out int indentLevel, out HorizontalAlignment horizontalAlignment, out var verticalAlignment, out var hidden, out var locked);

                add(new ExtendedFormat(xfId, -1, numFmtId, locked, hidden, indentLevel, horizontalAlignment, verticalAlignment));

                // reader.Skip();
            }
            else if (!XmlReaderHelper.SkipContent(reader))
            {
                break;
            }
        }

        static void ReadAlignment(XmlReader reader, string nsSpreadsheetMl, out int indentLevel, out HorizontalAlignment horizontalAlignment, out VerticalAlignment verticalAlignment, out bool hidden, out bool locked)
        {
            indentLevel = 0;
            horizontalAlignment = HorizontalAlignment.General;
            verticalAlignment = VerticalAlignment.Bottom;
            hidden = false;
            locked = false;

            if (!XmlReaderHelper.ReadFirstContent(reader))
            {
                return;
            }

            while (!reader.EOF)
            {
                if (reader.IsStartElement(NAlignment, nsSpreadsheetMl))
                {
                    int.TryParse(reader.GetAttribute(AIndent), NumberStyles.Integer, CultureInfo.InvariantCulture, out indentLevel);

                    var attrValue = reader.GetAttribute(AHorizontal);
                    if (attrValue is not null)
                    {
                        Enum.TryParse(attrValue, true, out horizontalAlignment);
                    }

                    attrValue = reader.GetAttribute(AVertical);
                    if (attrValue is not null)
                    {
                        Enum.TryParse(attrValue, true, out verticalAlignment);
                    }

                    reader.Skip();
                }
                else if (reader.IsStartElement(NProtection, nsSpreadsheetMl))
                {
                    locked = reader.GetAttribute(ALocked) == "1";
                    hidden = reader.GetAttribute(AHidden) == "1";
                    reader.Skip();
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }
        }
    }
}
