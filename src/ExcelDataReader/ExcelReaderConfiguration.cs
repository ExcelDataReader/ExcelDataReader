namespace ExcelDataReader
{
    /// <summary>
    /// Configuration options for an instance of ExcelDataReader.
    /// </summary>
    public class ExcelReaderConfiguration
    {
        /// <summary>
        /// Gets or sets a value indicating whether OLE Automation dates will be converted to DateTime. Default: true. (XLS only)
        /// </summary>
        public bool ConvertOaDate { get; set; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether BIFF parsing is Strict or Loose. Default: Strict. (XLS only)
        /// </summary>
        public ReadOption ReadOption { get; set; } = ReadOption.Strict;
    }
}
