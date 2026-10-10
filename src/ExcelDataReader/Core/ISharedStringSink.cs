namespace ExcelDataReader.Core;

/// <summary>
/// Receives shared strings as UTF-16LE code units borrowed from a parser buffer.
/// </summary>
internal interface ISharedStringSink
{
    /// <summary>
    /// Adds a string. The buffer is only valid for the duration of the call and must not be retained.
    /// </summary>
    void AddUtf16(byte[] buffer, int offset, int characterCount);
}
