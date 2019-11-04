﻿using System;

namespace ExcelDataReader.Core
{
    internal class Cell
    {
        public Cell(int columnIndex, object value, ExtendedFormat effectiveStyle)
        {
            ColumnIndex = columnIndex;
            Value = value;
            EffectiveStyle = effectiveStyle;
        }

        /// <summary>
        /// Gets the zero-based column index.
        /// </summary>
        public int ColumnIndex { get; }

        /// <summary>
        /// Gets the effective style on the cell. The effective style is determined from
        /// the Cell XF, with optional overrides from a Cell Style XF.
        /// </summary>
        public ExtendedFormat EffectiveStyle { get; }

        public object Value { get; }
    }
}
