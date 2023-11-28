namespace ExcelDataReader.Core
{
    internal sealed class ExtendedFormat
    {
        public ExtendedFormat(int numberFormatIndex)
        {
            NumberFormatIndex = numberFormatIndex;
        }

        public ExtendedFormat(int parentCellStyleXf, int fontIndex, int numberFormatIndex, bool locked, bool hidden, int indentLevel, HorizontalAlignment horizontalAlignment)
        {
            ParentCellStyleXf = parentCellStyleXf;
            FontIndex = fontIndex;
            NumberFormatIndex = numberFormatIndex;
            Locked = locked;
            Hidden = hidden;
            IndentLevel = indentLevel;
            HorizontalAlignment = horizontalAlignment;
        }

        private ExtendedFormat()
        {
        }

        public static ExtendedFormat Zero { get; } = new ExtendedFormat();

        /// <summary>
        /// Gets the to the parent Cell Style CF record with overrides for this XF. Only used with Cell XFs.
        /// 0xFFF means no override.
        /// </summary>
        public int ParentCellStyleXf { get; }

        public int FontIndex { get; }

        public int NumberFormatIndex { get; }

        public bool Locked { get; }

        public bool Hidden { get; }

        public int IndentLevel { get; }

        public HorizontalAlignment HorizontalAlignment { get; }
    }
}
