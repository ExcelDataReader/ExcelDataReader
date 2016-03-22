using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Core
{
   public class MemoryData
    {
        public static bool excelSavedFromAccessBool;

        public static List<List<String>> zipfilelist = new List<List<String>>();

        private static void zipfilelistadd(List<string> zipfileitems)
        {
            zipfilelist.Add(zipfileitems);
        }
        public static List<List<string>> returnlist
        {
            get { return zipfilelist; }
        }

    }
}
