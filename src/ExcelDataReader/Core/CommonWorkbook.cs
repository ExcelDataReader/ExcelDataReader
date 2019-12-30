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
            if (xfIndex >= 0 && xfIndex < ExtendedFormats.Count)
            {
                return ExtendedFormats[xfIndex];
            }

            return new ExtendedFormat()
            {
                NumberFormatIndex = numberFormatFromCell,
            };
        }

        /// <summary>
        /// Registers a number format string in the workbook's Formats dictionary.
        /// </summary>
        public void AddNumberFormat(int formatIndexInFile, string formatString)
        {
            if (!Formats.ContainsKey(formatIndexInFile))
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
