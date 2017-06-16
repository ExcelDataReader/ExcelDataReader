using System.Collections.Generic;

namespace ExcelDataReader.Core
{
    /// <summary>
    /// The common worksheet interface between the binary and OpenXml formats
    /// </summary>
    internal interface IWorksheet
    {
        string Name { get; }

        string VisibleState { get; }

        int FieldCount { get; }

        IEnumerable<object[]> ReadRows();
    }
}
