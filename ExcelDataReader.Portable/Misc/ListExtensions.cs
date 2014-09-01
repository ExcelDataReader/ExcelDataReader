using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExcelDataReader.Portable.Misc
{
    public static class ListExtensions
    {
        public static ReadOnlyCollection<T> AsReadOnly<T>(this List<T> list)
        {
            return new ReadOnlyCollection<T>(list);
        }
    }
}
