namespace ExcelDataReader.Portable.IO
{
    /// <summary>
    /// Implement this to provide implementations for your platform
    /// </summary>
    public interface IFileHelper
    {
        /// <summary>
        /// Returns path to a temporary folder
        /// e.g.
        /// System.IO.Path.GetTempPath();
        /// </summary>
        /// <returns></returns>
        string GetTempPath();
    }
}
