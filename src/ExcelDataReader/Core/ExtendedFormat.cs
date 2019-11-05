namespace ExcelDataReader.Core
{
    internal class ExtendedFormat
    {
        /// <summary>
        /// Gets or sets the index to the parent Cell Style CF record with overrides for this XF. Only used with Cell XFs.
        /// 0xFFF means no override
        /// </summary>
        public int ParentCellStyleXf { get; set; }

        public int FontIndex { get; set; }

        public int NumberFormatIndex { get; set; }

        public bool Locked { get; set; }

        public bool Hidden { get; set; }

        public int IndentLevel { get; set; }

        public HorizontalAlignment HorizontalAlignment { get; set; }
    }
}
