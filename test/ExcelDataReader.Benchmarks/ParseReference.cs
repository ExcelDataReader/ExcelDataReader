using System;
using System.Globalization;
using BenchmarkDotNet;
using BenchmarkDotNet.Attributes;

namespace ExcelDataReader.Benchmarks
{
    [MemoryDiagnoser]
    public class ParseReference
    {
        [Params("A1", "Z9", "XFD1048576")]
        public string CellReference { get; set; }

        [Benchmark]
        public (int, int) TryParseSubstring()
        {
            ReferenceHelper.ParseReferenceSubstring(CellReference, out int column, out int row);
            return (column, row);
        }

#if !NETFRAMEWORK
        [Benchmark]
        public (int, int) TryParseSpan()
        {
            ReferenceHelper.ParseReferenceSpan(CellReference, out int column, out int row);
            return (column, row);
        }
#endif

        [Benchmark]
        public (int, int) CustomTryParse()
        {
            ReferenceHelper.ParseReferenceCustom(CellReference, out int column, out int row);
            return (column, row);
        }

        private static class ReferenceHelper
        {
#if !NETFRAMEWORK
            public static bool ParseReferenceSpan(string value, out int column, out int row)
            {
                column = 0;
                var position = 0;
                const int offset = 'A' - 1;

                if (value != null)
                {
                    while (position < value.Length)
                    {
                        var c = value[position];
                        if (c >= 'A' && c <= 'Z')
                        {
                            position++;
                            column *= 26;
                            column += c - offset;
                            continue;
                        }

                        if (char.IsDigit(c))
                            break;

                        position = 0;
                        break;
                    }
                }

                if (position == 0)
                {
                    column = 0;
                    row = 0;
                    return false;
                }

                return int.TryParse(value.AsSpan(position), NumberStyles.None, CultureInfo.InvariantCulture, out row);
            }
#endif

            public static bool ParseReferenceSubstring(string value, out int column, out int row)
            {
                column = 0;
                var position = 0;
                const int offset = 'A' - 1;

                if (value != null)
                {
                    while (position < value.Length)
                    {
                        var c = value[position];
                        if (c >= 'A' && c <= 'Z')
                        {
                            position++;
                            column *= 26;
                            column += c - offset;
                            continue;
                        }

                        if (char.IsDigit(c))
                            break;

                        position = 0;
                        break;
                    }
                }

                if (position == 0)
                {
                    column = 0;
                    row = 0;
                    return false;
                }

                return int.TryParse(value.Substring(position), NumberStyles.None, CultureInfo.InvariantCulture, out row);
            }

            public static bool ParseReferenceCustom(string value, out int column, out int row)
            {
                column = 0;
                var position = 0;
                const int offset = 'A' - 1;

                if (value != null)
                {
                    while (position < value.Length)
                    {
                        var c = value[position];
                        if (c >= 'A' && c <= 'Z')
                        {
                            position++;
                            column *= 26;
                            column += c - offset;
                            continue;
                        }

                        if (char.IsDigit(c))
                            break;

                        position = 0;
                        break;
                    }
                }

                if (position == 0)
                {
                    column = 0;
                    row = 0;
                    return false;
                }

                return TryParseDecInt(value, position, out row);
            }


            private static bool IsDigit(int ch) => ((uint)ch - '0') <= 9;

            private static bool TryParseDecInt(string s, int startIndex, out int result)
            {
                if (startIndex >= s.Length)
                    goto Fail;

                long r = 0;
                for (int i = startIndex; i < s.Length; ++i)
                {
                    int num = s[i];
                    if (!IsDigit(num))
                        goto Fail;
                    r = r * 10 + (num - '0');
                    if (r > int.MaxValue)
                        goto Fail;
                }

                result = (int)r;
                return true;

                Fail:
                result = 0;
                return false;
            }
        }
    }
}
