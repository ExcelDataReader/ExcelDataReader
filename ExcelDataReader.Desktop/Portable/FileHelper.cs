using System;
using ExcelDataReader.Portable.IO;

namespace ExcelDataReader.Desktop.Portable
{
    public class FileHelper : IFileHelper
    {
        public string GetTempPath()
        {
            return System.IO.Path.GetTempPath();
        }
    }

}
