using System;

#nullable enable

namespace ExcelDataReader
{
    /// <summary>
    /// Header and footer text. 
    /// </summary>
    public sealed class HeaderFooter
    {
        internal HeaderFooter(bool hasDifferentFirst, bool hasDifferentOddEven)
        {
            HasDifferentFirst = hasDifferentFirst;
            HasDifferentOddEven = hasDifferentOddEven;
        }

        internal HeaderFooter(string? footer, string? header)
            : this(false, false)
        {
            OddHeader = header;
            OddFooter = footer;
        }

        /// <summary>
        /// Gets a value indicating whether the header and footer are different on the first page. 
        /// </summary>
        public bool HasDifferentFirst { get; }

        /// <summary>
        /// Gets a value indicating whether the header and footer are different on odd and even pages.
        /// </summary>
        public bool HasDifferentOddEven { get; }

        /// <summary>
        /// Gets the header used for the first page if <see cref="HasDifferentFirst"/> is <see langword="true"/>.
        /// </summary>
        public string? FirstHeader { get; internal set; }

        /// <summary>
        /// Gets the footer used for the first page if <see cref="HasDifferentFirst"/> is <see langword="true"/>.
        /// </summary>
        public string? FirstFooter { get; internal set; }

        /// <summary>
        /// Gets the header used for odd pages -or- all pages if <see cref="HasDifferentOddEven"/> is <see langword="false"/>. 
        /// </summary>
        public string? OddHeader { get; internal set; }

        /// <summary>
        /// Gets the footer used for odd pages -or- all pages if <see cref="HasDifferentOddEven"/> is <see langword="false"/>. 
        /// </summary>
        public string? OddFooter { get; internal set; }

        /// <summary>
        /// Gets the header used for even pages if <see cref="HasDifferentOddEven"/> is <see langword="true"/>. 
        /// </summary>
        public string? EvenHeader { get; internal set; }

        /// <summary>
        /// Gets the footer used for even pages if <see cref="HasDifferentOddEven"/> is <see langword="true"/>. 
        /// </summary>
        public string? EvenFooter { get; internal set; }
    }
}