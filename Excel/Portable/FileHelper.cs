using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader.IO;

namespace Excel.Portable
{
    public class FileHelper : IFileHelper
    {
        public string GetTempPath()
        {
            return System.IO.Path.GetTempPath();
        }
    }

}
