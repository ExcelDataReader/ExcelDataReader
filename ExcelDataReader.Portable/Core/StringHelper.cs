using System.Text;

namespace ExcelDataReader.Portable.Core
{
    public static class StringHelper
    {
        public static bool IsSingleByteEncoding(this Encoding encoding)
        {
            return encoding.GetByteCount(new char[] { 'a' }) == 1;
        }
    }
}
