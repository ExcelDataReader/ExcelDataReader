namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    /// <summary>
    /// Keeps the proper XML namespaces according to the format of the EXCEL file.
    /// </summary>
    internal static class XmlProperNamespaces
    {
        /// <summary>
        /// Gets the SpreadsheetMl namespace
        /// </summary>
        public static string NsSpreadsheetMl { get; private set; } = XmlNamespaces.NsSpreadsheetMl;

        /// <summary>
        /// Gets the DocumentRelationship namespace
        /// </summary>
        public static string NsDocumentRelationship { get; private set; } = XmlNamespaces.NsDocumentRelationship;

        /// <summary>
        /// Set the strict namespaces
        /// </summary>
        public static void SetStrictNamespaces()
        {
            NsSpreadsheetMl = XmlNamespaces.StrictNsSpreadsheetMl;
            NsDocumentRelationship = XmlNamespaces.StrictNsDocumentRelationship;
        }
    }
}