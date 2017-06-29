namespace ExcelDataReader.Core.OpenXmlFormat
{
    /// <summary>
    /// Base class for worksheet stream elements
    /// </summary>
    internal class XlsxElement
    {
        public XlsxElement(XlsxElementType type)
        {
            Type = type;
        }

        public XlsxElementType Type { get; }
    }
}
