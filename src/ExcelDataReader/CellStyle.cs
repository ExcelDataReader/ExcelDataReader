using System;

namespace ExcelDataReader
{
    /// <summary>
    /// Horizontal alignment. 
    /// </summary>
    public enum HorizontalAlignment
    {
        /// <summary>
        /// General.
        /// </summary>
        General,

        /// <summary>
        /// Left.
        /// </summary>
        Left,

        /// <summary>
        /// Centered.
        /// </summary>
        Centered,

        /// <summary>
        /// Right.
        /// </summary>
        Right,

        /// <summary>
        /// Filled.
        /// </summary>
        Filled,

        /// <summary>
        /// Justified.
        /// </summary>
        Justified,

        /// <summary>
        /// Centered across selection.
        /// </summary>
        CenteredAcrossSelection,

        /// <summary>
        /// Distributed.
        /// </summary>
        Distributed,
    }

    /// <summary>
    /// Holds style information for a cell.
    /// </summary>
    public class CellStyle
    {
        /// <summary>
        /// Gets the font index.
        /// </summary>
        public int FontIndex { get; internal set; }

        /// <summary>
        /// Gets the number format index.
        /// </summary>
        public int NumberFormatIndex { get; internal set; }

        /// <summary>
        /// Gets the indent level.
        /// </summary>
        public int IndentLevel { get; internal set; }

        /// <summary>
        /// Gets the horizontal alignment.
        /// </summary>
        public HorizontalAlignment HorizontalAlignment { get; internal set; }

        /// <summary>
        /// Gets a value indicating whether the cell is hidden.
        /// </summary>
        public bool Hidden { get; internal set; }

        /// <summary>
        /// Gets a value indicating whether the cell is locked.
        /// </summary>
        public bool Locked { get; internal set; }
    }
}