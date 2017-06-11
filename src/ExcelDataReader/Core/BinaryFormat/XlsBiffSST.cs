using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a Shared String Table in BIFF8 format
    /// </summary>
    internal class XlsBiffSST : XlsBiffRecord
    {
        private readonly List<uint> _continues = new List<uint>();
        private readonly List<string> _strings;

        internal XlsBiffSST(byte[] bytes, uint offset, bool isV8, Encoding encoding)
            : base(bytes, offset)
        {
            _strings = new List<string>();
            IsV8 = isV8;
            SSTEncoding = encoding;
        }

        /// <summary>
        /// Gets the number of strings in SST
        /// </summary>
        public uint Count => ReadUInt32(0x0);

        /// <summary>
        /// Gets the count of unique strings in SST
        /// </summary>
        public uint UniqueCount => ReadUInt32(0x4);

        public bool IsV8 { get; }

        public Encoding SSTEncoding { get; }

        /// <summary>
        /// Reads strings from BIFF stream into SST array
        /// </summary>
        public void ReadStrings()
        {
            uint offset = (uint)RecordContentOffset + 8;
            uint last = (uint)RecordContentOffset + RecordSize;
            int lastcontinue = 0;
            uint count = UniqueCount;
            while (offset < last)
            {
                var str = XlsStringFactory.CreateXlsString(Bytes, offset, IsV8, SSTEncoding);
                uint prefix = str.HeadSize;
                uint postfix = str.TailSize;
                uint len = str.CharacterCount;
                uint size = prefix + postfix + len + (str.IsMultiByte ? len : 0);
                if (offset + size > last)
                {
                    if (lastcontinue >= _continues.Count)
                        break;
                    uint contoffset = _continues[lastcontinue];

                    byte[] buff = new byte[size * 2];
                    Buffer.BlockCopy(Bytes, (int)offset, buff, 0, (int)(last - offset));

                    // If we're past the string data then we won't have a unicode string option flags.
                    if (offset + prefix + len + (str.IsMultiByte ? len : 0) <= last)
                    {
                        Buffer.BlockCopy(Bytes, (int)contoffset + 4, buff, (int)(last - offset), (int)(size - last + offset));
                        offset = contoffset + 4 + size - last + offset;
                    }
                    else
                    {
                        bool isMultiByte = (Buffer.GetByte(Bytes, (int)contoffset + 4) & 0x1) == 1;
                        if (!isMultiByte && str.IsMultiByte)
                        {
                            len -= (last - prefix - offset) / 2;
                            byte[] tempbytes = new byte[len * 2];
                            for (int i = 0; i < len; i++)
                            {
                                tempbytes[i * 2] = Bytes[contoffset + 5 + i];
                            }

                            Buffer.BlockCopy(tempbytes, 0, buff, (int)(last - offset), tempbytes.Length);
                            Buffer.BlockCopy(Bytes, (int)(contoffset + 5 + len), buff, (int)(last - offset + tempbytes.Length), (int)postfix);
                            offset = contoffset + 5 + len + postfix;
                        }
                        else if (isMultiByte && !str.IsMultiByte)
                        {
                            len -= last - offset - prefix;

                            int templen = (int)(last - offset - prefix);
                            byte[] tempbytes = new byte[templen * 2];
                            for (int i = 0; i < templen; i++)
                            {
                                tempbytes[i * 2] = Bytes[offset + prefix + i];
                            }

                            Buffer.BlockCopy(tempbytes, 0, buff, (int)prefix, tempbytes.Length);
                            int buffOffset = (int)(prefix + tempbytes.Length);

                            Buffer.BlockCopy(Bytes, (int)(contoffset + 5), buff, buffOffset, (int)(len + len));
                            Buffer.BlockCopy(Bytes, (int)(contoffset + 5 + len + len), buff, (int)(buffOffset + len + len), (int)postfix);
                            buff[2] = (byte)((XlsFormattedUnicodeString.FormattedUnicodeStringFlags)buff[2] | XlsFormattedUnicodeString.FormattedUnicodeStringFlags.MultiByte);
                            offset = contoffset + 5 + len + len + postfix;
                        }
                        else
                        {
                            Buffer.BlockCopy(Bytes, (int)contoffset + 5, buff, (int)(last - offset), (int)(size - last + offset));
                            offset = contoffset + 5 + size - last + offset;
                        }
                    }

                    last = contoffset + 4 + BitConverter.ToUInt16(Bytes, (int)contoffset + 2);
                    lastcontinue++;

                    str = new XlsFormattedUnicodeString(buff, 0);
                }
                else
                {
                    offset += size;
                    if (offset == last)
                    {
                        if (lastcontinue < _continues.Count)
                        {
                            uint contoffset = _continues[lastcontinue];
                            offset = contoffset + 4;
                            last = offset + BitConverter.ToUInt16(Bytes, (int)contoffset + 2);
                            lastcontinue++;
                        }
                        else
                        {
                            count = 1;
                        }
                    }
                }

                _strings.Add(str.Value);
                count--;
                if (count == 0)
                    break;
            }
        }

        /// <summary>
        /// Returns string at specified index
        /// </summary>
        /// <param name="sstIndex">Index of string to get</param>
        /// <returns>string value if it was found, empty string otherwise</returns>
        public string GetString(uint sstIndex)
        {
            if (sstIndex < _strings.Count)
                return _strings[(int)sstIndex];
            
            return string.Empty;
        }

        /// <summary>
        /// Appends Continue record to SST
        /// </summary>
        /// <param name="fragment">Continue record</param>
        public void Append(XlsBiffContinue fragment)
        {
            _continues.Add((uint)fragment.Offset);
        }
    }
}