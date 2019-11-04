using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal abstract class Record
    {
        internal static Record Default { get; } = new DefaultRecord();

        private sealed class DefaultRecord : Record
        {
        }
    }
}
