using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using System.IO;
using System.Data;

namespace MichaelsExcelTester
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"\\filestore\Teams\Professional Services\McDONALD'S\Vendor Management\ABFS\2015 BB Upgrade Worklist_DR.xlsx";
            using(FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using(MemoryStream ms = new MemoryStream())
                {
                    fs.CopyTo(ms);
                    Console.WriteLine("loaded into mem");
                    var reader = ExcelReaderFactory.CreateOpenXmlReader(ms);
                    Console.WriteLine("reader initialized");
                    DataSet ds = reader.AsDataSet();
                    Console.WriteLine(string.Join(",", ds.Tables));
                }
            }
            
            
        }
    }
}
