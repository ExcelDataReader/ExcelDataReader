using System;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    /// <summary>
    /// Keeps the proper XML namespaces according to the format of the EXCEL file.
    /// </summary>
    internal sealed class XmlProperNamespaces
    {
        public XmlProperNamespaces(bool isStrict)
        {
            NsSpreadsheetMl = isStrict ? XmlNamespaces.StrictNsSpreadsheetMl : XmlNamespaces.NsSpreadsheetMl;
            NsDocumentRelationship = isStrict ? XmlNamespaces.StrictNsDocumentRelationship : XmlNamespaces.NsDocumentRelationship;
        }

        /// <summary>
        /// Gets the SpreadsheetMl namespace. 
        /// </summary>
        public string NsSpreadsheetMl { get; }

        /// <summary>
        /// Gets the DocumentRelationship namespace.
        /// </summary>
        public string NsDocumentRelationship { get; }
    }
}