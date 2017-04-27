using System;
using System.Collections.Generic;
using System.Text;
using Excel;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents a Shared String Table in BIFF8 format
    /// </summary>
    internal class XlsBiffSST : XlsBiffRecord
    {
        private readonly List<uint> m_continues = new List<uint>();
        private readonly List<string> m_strings;

        internal XlsBiffSST(byte[] bytes, uint offset, ExcelBinaryReader reader)
            : base(bytes, offset, reader)
        {
            m_strings = new List<string>();
        }

        /// <summary>
        /// Returns count of strings in SST
        /// </summary>
        public uint Count => ReadUInt32(0x0);

        /// <summary>
        /// Returns count of unique strings in SST
        /// </summary>
        public uint UniqueCount => ReadUInt32(0x4);

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
                var str = XlsStringFactory.CreateXlsString(Bytes, offset, Reader);
                uint prefix = str.HeadSize;
                uint postfix = str.TailSize;
                uint len = str.CharacterCount;
                uint size = prefix + postfix + len + (str.IsMultiByte ? len : 0);
                if (offset + size > last)
                {
                    if (lastcontinue >= m_continues.Count)
                        break;
                    uint contoffset = m_continues[lastcontinue];

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
                            len -= (last - offset - prefix);

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
                        if (lastcontinue < m_continues.Count)
                        {
                            uint contoffset = m_continues[lastcontinue];
                            offset = contoffset + 4;
                            last = offset + BitConverter.ToUInt16(Bytes, (int)contoffset + 2);
                            lastcontinue++;
                        }
                        else
                            count = 1;
                    }
                }
                m_strings.Add(str.Value);
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
            if (sstIndex < m_strings.Count)
                return m_strings[(int)sstIndex];


            return string.Empty;
        }

        /// <summary>
        /// Appends Continue record to SST
        /// </summary>
        /// <param name="fragment">Continue record</param>
        public void Append(XlsBiffContinue fragment)
        {
            m_continues.Add((uint)fragment.Offset);
        }
    }
}