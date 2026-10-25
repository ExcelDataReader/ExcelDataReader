using System.Text;

namespace ExcelDataReader.Core.CompoundFormat;

internal static class CompoundDocumentPropertiesReader
{
    private static readonly Guid SummaryInformation = new("F29F85E0-4FF9-1068-AB91-08002B27B3D9");

    private static long MaximumFileTime { get; } = DateTime.MaxValue.ToFileTimeUtc();

    public static DocumentProperties Load(CompoundDocument document, Stream source)
    {
        var entry = document.FindEntry("\u0005SummaryInformation");
        if (entry == null || entry.EntryType != STGTY.STGTY_STREAM || entry.StreamSize < 28 || entry.StreamSize > int.MaxValue)
            return DocumentProperties.Empty;

        long position = source.Position;
        Exception? failure = null;
        try
        {
            using var stream = new CompoundStream(document, source, entry.StreamFirstSector, (int)entry.StreamSize, entry.IsEntryMiniStream, true);
            return Read(stream);
        }
        catch (Exception error)
        {
            failure = error;
            throw;
        }
        finally
        {
            ResourceCleanup.RestorePosition(source, position, failure);
        }
    }

    internal static DocumentProperties Read(Stream stream)
    {
        var properties = new DocumentPropertiesBuilder();
        using var reader = new BinaryReader(stream, Encoding.Unicode, leaveOpen: true);
        if (stream.Length < 28 || reader.ReadUInt16() != 0xFFFE)
            return DocumentProperties.Empty;
        stream.Position = 24;
        uint sections = reader.ReadUInt32();
        if (sections > (stream.Length - 28) / 20)
            return DocumentProperties.Empty;

        for (uint i = 0; i < sections; i++)
        {
            stream.Position = 28 + i * 20L;
            var format = new Guid(reader.ReadBytes(16));
            uint offset = reader.ReadUInt32();
            if (format != SummaryInformation || offset < 28 + sections * 20L || offset > stream.Length - 8)
                continue;
            stream.Position = offset;
            uint size = reader.ReadUInt32();
            uint count = reader.ReadUInt32();
            if (size < 8 || size > stream.Length - offset || count > (size - 8) / 8)
                continue;

            var section = new PropertySection(reader, offset, size, count);
            Encoding? encoding = null;
            if (section.Seek(1, out long end) && stream.Position + 6 <= end && reader.ReadUInt32() == 2)
                encoding = GetEncoding(reader.ReadUInt16());
            for (uint id = 2; id <= 13; id++)
            {
                if (id is 7 or 9 or 10 || !section.Seek(id, out end) || stream.Position + 4 > end)
                    continue;
                uint type = reader.ReadUInt32();
                if (id >= 11)
                {
                    if (type != 64 || stream.Position + 8 > end)
                        continue;
                    long value = reader.ReadInt64();
                    if (value < 0 || value > MaximumFileTime)
                        continue;
                    var date = DateTime.FromFileTimeUtc(value);
                    switch (id)
                    {
                        case 11: properties.LastPrinted = date; break;
                        case 12: properties.Created = date; break;
                        case 13: properties.Modified = date; break;
                    }
                }
                else
                {
                    string? value = ReadString(reader, type, encoding, end);
                    if (value == null)
                        continue;
                    switch (id)
                    {
                        case 2: properties.Title = value; break;
                        case 3: properties.Subject = value; break;
                        case 4: properties.Author = value; break;
                        case 5: properties.Keywords = value; break;
                        case 6: properties.Description = value; break;
                        case 8: properties.LastModifiedBy = value; break;
                    }
                }
            }
        }

        return properties.Build();
    }

    private static string? ReadString(BinaryReader reader, uint type, Encoding? encoding, long end)
    {
        var stream = reader.BaseStream;
        if (type is not (30 or 31) || stream.Position + 4 > end)
            return null;
        uint count = reader.ReadUInt32();
        int width = type == 31 ? 2 : 1;
        if (count > (end - stream.Position) / width || count > int.MaxValue / width)
            return null;
        var decoder = type == 31 ? Encoding.Unicode : encoding;
        if (decoder == null)
            return null;

        byte[] bytes = reader.ReadBytes((int)count * width);
        if (bytes.Length != (int)count * width)
            throw new EndOfStreamException();
        string value = decoder.GetString(bytes);
        int terminator = value.IndexOf('\0');
        return terminator >= 0 ? value.Substring(0, terminator) : value;
    }

    private static Encoding? GetEncoding(ushort codePage)
    {
        try
        {
            return EncodingHelper.GetEncoding(codePage);
        }
        catch (ArgumentException)
        {
            // Unsupported property encodings do not affect Unicode properties or cells.
            return null;
        }
        catch (NotSupportedException)
        {
            return null;
        }
    }

    private sealed class PropertySection(BinaryReader reader, uint offset, uint size, uint count)
    {
        public bool Seek(uint id, out long end)
        {
            uint valueOffset = 0;
            for (uint i = 0; i < count; i++)
            {
                reader.BaseStream.Position = offset + 8 + i * 8L;
                uint property = reader.ReadUInt32();
                uint location = reader.ReadUInt32();
                if (property == id)
                    valueOffset = location;
            }

            end = offset + (long)size;
            if (valueOffset < 8 + count * 8L || valueOffset >= size)
                return false;

            for (uint i = 0; i < count; i++)
            {
                reader.BaseStream.Position = offset + 12 + i * 8L;
                uint location = reader.ReadUInt32();
                if (location > valueOffset && location < size)
                    end = Math.Min(end, offset + (long)location);
            }

            reader.BaseStream.Position = offset + (long)valueOffset;
            return true;
        }
    }
}
