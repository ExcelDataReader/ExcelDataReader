using System;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    /// <summary>
    /// Keeps the proper XML namespaces according to the format of the EXCEL file.
    /// </summary>
    internal class XmlProperNamespaces
    {
        /// <summary>
        /// Gets the SpreadsheetMl namespace
        /// </summary>
        public string NsSpreadsheetMl { get; private set; } = XmlNamespaces.NsSpreadsheetMl;

        /// <summary>
        /// Gets the DocumentRelationship namespace
        /// </summary>
        public string NsDocumentRelationship { get; private set; } = XmlNamespaces.NsDocumentRelationship;

        internal void SetStrictNamespaces()
        {
            NsSpreadsheetMl = XmlNamespaces.StrictNsSpreadsheetMl;
            NsDocumentRelationship = XmlNamespaces.StrictNsDocumentRelationship;
        }
    }
}