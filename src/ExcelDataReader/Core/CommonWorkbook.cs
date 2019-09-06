using System;
using System.Collections.Generic;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core
{
    /// <summary>
    /// Common handling of extended formats (XF) and mappings between file-based and global number format indices.
    /// </summary>
    internal class CommonWorkbook
    {
        public CommonWorkbook()
        {
            const int maxBuiltInFormats = 163;
            for (var i = 0; i < maxBuiltInFormats; i++)
            {
                var numFmt = BuiltinNumberFormat.GetBuiltinNumberFormat(i);
                if (numFmt != null)
                {
                    Formats.Add(i, numFmt);
                }
            }
        }

        /// <summary>
        /// Gets the dictionary of global number format strings. Always includes the built-in formats at their
        /// corresponding indices and any additional formats specified in the workbook file.
        /// </summary>
        public Dictionary<int, NumberFormatString> Formats { get; } = new Dictionary<int, NumberFormatString>();

        /// <summary>
        /// Gets the Cell XFs
        /// </summary>
        public List<ExtendedFormat> ExtendedFormats { get; } = new List<ExtendedFormat>();

        /// <summary>
        /// Gets the Cell Style XFs
        /// </summary>
        public List<ExtendedFormat> CellStyleExtendedFormats { get; } = new List<ExtendedFormat>();

        /// <summary>
        /// Gets the the dictionary of mappings between format index in the file and key in the Formats dictionary.
        /// </summary>
        private Dictionary<int, int> FormatMappings { get; } = new Dictionary<int, int>();

        public ExtendedFormat GetEffectiveCellStyle(int xfIndex, int numberFormatFromCell)
        {
            var effectiveStyle = new ExtendedFormat();
            var cellXf = xfIndex >= 0 && xfIndex < ExtendedFormats.Count
                ? ExtendedFormats[xfIndex]
                : null;
            if (cellXf != null)
            {
                effectiveStyle.FontIndex = cellXf.FontIndex;
                effectiveStyle.FormatIndex = GetNumberFormatFromFileIndex(cellXf.FormatIndex); // fileindex->

                effectiveStyle.Hidden = cellXf.Hidden;
                effectiveStyle.Locked = cellXf.Locked;
                effectiveStyle.IndentLevel = cellXf.IndentLevel;
                effectiveStyle.HorizontalAlignment = cellXf.HorizontalAlignment;

                var cellStyleXf = cellXf.ParentCellStyleXf >= 0 && cellXf.ParentCellStyleXf < CellStyleExtendedFormats.Count 
                    ? CellStyleExtendedFormats[cellXf.ParentCellStyleXf] 
                    : null;
                if (cellStyleXf != null)
                {
                    if (cellStyleXf.ApplyFont)
                    {
                        effectiveStyle.FontIndex = cellStyleXf.FontIndex;
                    }

                    if (cellStyleXf.ApplyNumberFormat)
                    {
                        effectiveStyle.FormatIndex = GetNumberFormatFromFileIndex(cellStyleXf.FormatIndex);
                    }

                    if (cellStyleXf.ApplyProtection)
                    {
                        effectiveStyle.Hidden = cellStyleXf.Hidden;
                        effectiveStyle.Locked = cellStyleXf.Locked;
                    }

                    if (cellStyleXf.ApplyTextAlignment)
                    {
                        effectiveStyle.IndentLevel = cellStyleXf.IndentLevel;
                        effectiveStyle.HorizontalAlignment = cellStyleXf.HorizontalAlignment;
                    }
                }
            }
            else
            {
                effectiveStyle.FormatIndex = GetNumberFormatFromFileIndex(numberFormatFromCell);
            }

            return effectiveStyle;
        }

        /// <summary>
        /// Returns the global number format index from a file-based format index.
        /// </summary>
        public int GetNumberFormatFromFileIndex(int formatIndexInFile)
        {
            if (FormatMappings.TryGetValue(formatIndexInFile, out var formatIndex))
            {
                return formatIndex;
            }

            // Format not stored in file, assume built-in format
            return formatIndexInFile;
        }

        /// <summary>
        /// Registers a number format string and its file-based format index in the workbook's Formats dictionary.
        /// If the format string matches a built-in or previously registered format, it will be mapped to that index.
        /// </summary>
        public void AddNumberFormat(int formatIndexInFile, string formatString)
        {
            var exists = false;
            int maxIndex = 163;
            foreach (var format in Formats)
            {
                if (!exists && format.Value.FormatString == formatString)
                {
                    FormatMappings[formatIndexInFile] = format.Key;
                    exists = true;
                }

                maxIndex = Math.Max(maxIndex, format.Key);
            }

            if (!exists)
            {
                maxIndex++;
                Formats.Add(maxIndex, new NumberFormatString(formatString));
                FormatMappings[formatIndexInFile] = maxIndex;
            }
        }
    }
}
