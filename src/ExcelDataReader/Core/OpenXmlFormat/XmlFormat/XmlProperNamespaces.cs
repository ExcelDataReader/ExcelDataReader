namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

/// <summary>
/// Keeps the proper XML namespaces according to the format of the EXCEL file.
/// </summary>
internal sealed class XmlProperNamespaces(bool isStrict)
{
    /// <summary>
    /// Gets the SpreadsheetMl namespace. 
    /// </summary>
    public string NsSpreadsheetMl { get; } = isStrict ? XmlNamespaces.StrictNsSpreadsheetMl : XmlNamespaces.NsSpreadsheetMl;

    /// <summary>
    /// Gets the DocumentRelationship namespace.
    /// </summary>
    public string NsDocumentRelationship { get; } = isStrict ? XmlNamespaces.StrictNsDocumentRelationship : XmlNamespaces.NsDocumentRelationship;
}