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

        private NumberFormatString GeneralNumberFormat { get; } = new NumberFormatString("General");

        public ExtendedFormat GetEffectiveCellStyle(int xfIndex, int numberFormatFromCell)
        {
            var effectiveStyle = new ExtendedFormat();
            var cellXf = xfIndex >= 0 && xfIndex < ExtendedFormats.Count
                ? ExtendedFormats[xfIndex]
                : null;
            if (cellXf != null)
            {
                effectiveStyle.FontIndex = cellXf.FontIndex;
                effectiveStyle.NumberFormatIndex = cellXf.NumberFormatIndex;

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
                        effectiveStyle.NumberFormatIndex = cellStyleXf.NumberFormatIndex;
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
                effectiveStyle.NumberFormatIndex = numberFormatFromCell;
            }

            return effectiveStyle;
        }

        /// <summary>
        /// Registers a number format string in the workbook's Formats dictionary.
        /// </summary>
        public void AddNumberFormat(int formatIndexInFile, string formatString)
        {
            Formats.Add(formatIndexInFile, new NumberFormatString(formatString));
        }

        public NumberFormatString GetNumberFormatString(int numberFormatIndex)
        {
            if (Formats.TryGetValue(numberFormatIndex, out var numberFormat))
            {
                return numberFormat;
            }

            numberFormat = BuiltinNumberFormat.GetBuiltinNumberFormat(numberFormatIndex);
            if (numberFormat != null)
            {
                return numberFormat;
            }

            // Fall back to "General" if the number format index is invalid
            return GeneralNumberFormat;
        }
    }
}
