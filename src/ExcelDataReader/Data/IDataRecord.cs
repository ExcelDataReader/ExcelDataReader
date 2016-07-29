// Type: System.Data.IDataRecord
// Assembly: System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089
// Assembly location: C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Data.dll

using System;

// ReSharper disable CheckNamespace
namespace ExcelDataReader.Data
// ReSharper restore CheckNamespace
{
    public interface IDataRecord
    {
        int FieldCount { get; }

        object this[int i] { get; }

        object this[string name] { get; }

        string GetName(int i);

        string GetDataTypeName(int i);

        Type GetFieldType(int i);

        object GetValue(int i);

        int GetValues(object[] values);

        int GetOrdinal(string name);

        bool GetBoolean(int i);

        byte GetByte(int i);

        long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length);

        char GetChar(int i);

        long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length);

        Guid GetGuid(int i);

        short GetInt16(int i);

        int GetInt32(int i);

        long GetInt64(int i);

        float GetFloat(int i);

        double GetDouble(int i);

        string GetString(int i);

        Decimal GetDecimal(int i);

        DateTime GetDateTime(int i);

        //IDataReader GetData(int i);

        bool IsDBNull(int i);
    }
}
