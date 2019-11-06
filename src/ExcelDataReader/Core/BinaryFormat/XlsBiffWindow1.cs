using System;

namespace ExcelDataReader.Core.BinaryFormat
{
    /// <summary>
    /// Represents Workbook's global window description
    /// </summary>
    internal class XlsBiffWindow1 : XlsBiffRecord
    {
        internal XlsBiffWindow1(byte[] bytes)
            : base(bytes)
        {
        }

        [Flags]
        public enum Window1Flags : ushort
        {
            Hidden = 0x1,
            Minimized = 0x2,
            
            // (Reserved) = 0x4,
            HScrollVisible = 0x8,
            VScrollVisible = 0x10,
            WorkbookTabs = 0x20
            
            // (Other bits are reserved)
        }

        /// <summary>
        /// Gets the X position of a window
        /// </summary>
        public ushort Left => ReadUInt16(0x0);

        /// <summary>
        /// Gets the Y position of a window
        /// </summary>
        public ushort Top => ReadUInt16(0x2);

        /// <summary>
        /// Gets the width of the window
        /// </summary>
        public ushort Width => ReadUInt16(0x4);

        /// <summary>
        /// Gets the height of the window
        /// </summary>
        public ushort Height => ReadUInt16(0x6);

        /// <summary>
        /// Gets the window flags
        /// </summary>
        public Window1Flags Flags => (Window1Flags)ReadUInt16(0x8);

        /// <summary>
        /// Gets the active workbook tab (zero-based)
        /// </summary>
        public ushort ActiveTab => ReadUInt16(0xA);

        /// <summary>
        /// Gets the first visible workbook tab (zero-based)
        /// </summary>
        public ushort FirstVisibleTab => ReadUInt16(0xC);

        /// <summary>
        /// Gets the number of selected workbook tabs
        /// </summary>
        public ushort SelectedTabCount => ReadUInt16(0xE);

        /// <summary>
        /// Gets the workbook tab width to horizontal scrollbar width
        /// </summary>
        public ushort TabRatio => ReadUInt16(0x10);
    }
}