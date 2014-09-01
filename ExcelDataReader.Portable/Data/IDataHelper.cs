using System;

namespace ExcelDataReader.Portable.Data
{
    /// <summary>
    /// Implement this to provide implementations for your platform
    /// </summary>
    public interface IDataHelper
    {
// ReSharper disable InconsistentNaming
        bool IsDBNull(Object value);
// ReSharper restore InconsistentNaming
    }
}
