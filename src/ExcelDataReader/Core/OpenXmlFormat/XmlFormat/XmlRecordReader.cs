using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

internal abstract class XmlRecordReader(XmlReader reader) : RecordReader
{
    private IEnumerator<Record> _enumerator;

    public XmlProperNamespaces ProperNamespaces { get; set; } = new(reader.IsStartElement() && reader.NamespaceURI == XmlNamespaces.StrictNsSpreadsheetMl);

    protected XmlReader Reader { get; } = reader;

    public override Record Read()
    {
        _enumerator ??= ReadOverride().GetEnumerator();
        if (_enumerator.MoveNext())
            return _enumerator.Current;
        return null;
    }

    protected abstract IEnumerable<Record> ReadOverride();

    /// <inheritdoc />
    protected override void Dispose(bool disposing)
    {
        _enumerator?.Dispose();
        if (disposing)
            Reader.Dispose();
    }
}
