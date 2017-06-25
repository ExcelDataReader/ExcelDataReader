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
    }
}
