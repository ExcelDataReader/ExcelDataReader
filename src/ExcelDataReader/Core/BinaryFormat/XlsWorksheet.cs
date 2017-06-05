namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Worksheet section in workbook
    /// </summary>
    internal class XlsWorksheet
    {
        public XlsWorksheet(int index, XlsBiffBoundSheet refSheet)
        {
            Index = index;
            Name = refSheet.SheetName;
            DataOffset = refSheet.StartOffset;

            switch (refSheet.VisibleState)
            {
                case XlsBiffBoundSheet.SheetVisibility.Hidden:
                    VisibleState = "hidden";
                    break;
                case XlsBiffBoundSheet.SheetVisibility.VeryHidden:
                    VisibleState = "veryhidden";
                    break;
                default:
                    VisibleState = "visible";
                    break;
            }
        }

        /// <summary>
        /// Gets the worksheet name
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the visibility of worksheet
        /// </summary>
        public string VisibleState { get; }

        /// <summary>
        /// Gets the zero-based index of worksheet
        /// </summary>
        public int Index { get; }

        /// <summary>
        /// Gets the worksheet data offset.
        /// </summary>
        public uint DataOffset { get; }

        public XlsBiffSimpleValueRecord CalcMode { get; set; }

        public XlsBiffSimpleValueRecord CalcCount { get; set; }

        public XlsBiffSimpleValueRecord RefMode { get; set; }

        public XlsBiffSimpleValueRecord Iteration { get; set; }

        public XlsBiffRecord Delta { get; set; }
        
        public XlsBiffRecord Window { get; set; }
    }
}