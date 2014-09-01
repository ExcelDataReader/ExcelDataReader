using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader.Portable.IO;

namespace Excel.Portable
{
    public class FileHelper : IFileHelper
    {
        [Obsolete("no longer needed")]
        public string GetTempPath()
        {
            return System.IO.Path.GetTempPath();
        }
    }

}
